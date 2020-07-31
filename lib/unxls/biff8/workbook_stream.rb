# frozen_string_literal: true

class Unxls::Biff8::WorkbookStream

  attr_reader :parser

  # @param parser [Unxls::Parser]
  def initialize(parser)
    @parser = parser
    @parsed_result = []
    @cell_index = { dimensions: {}, hlinks: {}, hlinktooltips: {}, notes: {} }
    @stream = Ole::Storage.open(parser.file.tap(&:rewind)).file.open('Workbook')
  end

  # @return [Array<Hash>]
  def parse
    @stream.rewind
    decrypt

    @stream.rewind
    while (record = get_next_record)
      parsed_record = record.process

      case record.name
      when :BOF
        @parsed_result << { BOF: parsed_record }

      when :EOF
        @parsed_result.last[:EOF] = { EOF: parsed_record }

      else
        next unless parsed_record

        if record.serial?
          @parsed_result.last[record.name] ||= []
          @parsed_result.last[record.name] << parsed_record
        else
          @parsed_result.last[record.name] = parsed_record
        end

        update_address_index(parsed_record)
      end
    end

    @parsed_result << @cell_index
    @parsed_result.freeze

  ensure
    @stream.unlink if @stream.is_a?(Tempfile) # remember to delete the tempfile used for decrypted data
  end
  
  # @param sheet_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @return [Symbol]
  def self.make_address_key(sheet_index, row, column)
    "#{sheet_index}_#{row}_#{column}".to_sym
  end

  INDEXABLE_RECORDS = [
    :MulBlank,
    :Blank,
    :LabelSst,
    :MulRk,
    :Number,
    :RK,
    :Formula,
    :BoolErr,
    :HLink,
    :HLinkTooltip,
    :Note,
  ].freeze

  # Index for quick cell lookup
  # {
  #   :<sheet index>_<cell row>_<cell column> => :<record name>_<record index>
  # }
  # @param parsed_record [Hash]
  # @return [true]
  def update_address_index(parsed_record)
    return unless INDEXABLE_RECORDS.include?(parsed_record[:_record][:name])

    record_name = parsed_record[:_record][:name]
    sheet_index = @parsed_result.size - 1
    record_array_index = @parsed_result.last[record_name].size - 1
    access_address = "#{record_name}_#{record_array_index}".to_sym
    row = parsed_record[:rw]
    column = parsed_record[:col]
    
    case record_name
    when :MulBlank, :MulRk
      (parsed_record[:colFirst]..parsed_record[:colLast]).each do |column|
        sheet_row_col = self.class.make_address_key(sheet_index, row, column)
        @cell_index[sheet_row_col] = access_address
        update_value_dimensions(record_name, sheet_index, row, column)
      end
      
    when :HLink
      (parsed_record[:rwFirst]..parsed_record[:rwLast]).each do |row|
        (parsed_record[:colFirst]..parsed_record[:colLast]).each do |column|
          sheet_row_col = self.class.make_address_key(sheet_index, row, column)
          @cell_index[:hlinks][sheet_row_col] = record_array_index
        end
      end

    when :HLinkTooltip
      dimensions = parsed_record[:frtRefHeaderNoGrbit][:ref8]
      (dimensions[:rwFirst]..dimensions[:rwLast]).each do |row|
        (dimensions[:colFirst]..dimensions[:colLast]).each do |column|
          sheet_row_col = self.class.make_address_key(sheet_index, row, column)
          @cell_index[:hlinktooltips][sheet_row_col] = record_array_index
        end
      end
    
    when :Note
      sheet_row_col = self.class.make_address_key(sheet_index, row, column)
      @cell_index[:notes][sheet_row_col] = record_array_index
    
    else # :LabelSst, :RK, :Blank, :BoolErr, :Number, :Formula
      sheet_row_col = self.class.make_address_key(sheet_index, row, column)
      @cell_index[sheet_row_col] = access_address
      update_value_dimensions(record_name, sheet_index, row, column)
    end

    true
  end

  # Make sheet dimension tables using only the cells that contain a value,
  # unlike the Dimensions record, which honors empty formatted cells.
  # {
  #   :<sheet index> => { rmin: <top row>, rmax: <bottom row>, cmin: <left column>, cmax: <right column> },
  #   …
  # }
  # @param record_name [Symbol]
  # @param sheet_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @return [Hash]
  def update_value_dimensions(record_name, sheet_index, row, column)
    return unless %i(LabelSst MulRk Number RK Formula BoolErr).include?(record_name)

    dim = @cell_index[:dimensions]
    dim[sheet_index] ||= { rmin: row, rmax: row, cmin: column, cmax: column }

    dim = dim[sheet_index]
    dim[:rmin] = row if row < dim[:rmin]
    dim[:rmax] = row if row > dim[:rmax]
    dim[:cmin] = column if column < dim[:cmin]
    dim[:cmax] = column if column > dim[:cmax]
  end

  # @return [StringIO]
  def decrypt
    while (record = get_next_record)
      return if %i(InterfaceHdr WriteAccess CodePage).include?(record.name) # first mandatory records after the optional FilePass record
      break if record.name == :FilePass # stream is encoded if this record is present
    end

    password = @parser.settings[:password] || Unxls::Offcrypto::DEFAULT_PASSWORD
    decrypted_stream = Tempfile.new.tap(&:binmode) # remember to unlink the tempfile!
    filepass_data = record.process

    case (decryption_type = filepass_data[:_type])

    when :XOR
      ok = Unxls::Offcrypto._xor_password_match?(password, filepass_data[:verificationBytes])
      raise("Password '#{password}' does not match") unless ok

      xor_array = Unxls::Offcrypto._create_xor_array_method1(password)

      @stream.rewind
      until @stream.eof?
        pos = @stream.pos
        _, size = read_record_header

        @stream.seek(pos)
        decrypted_stream << @stream.read(Unxls::Biff8::Record::HEADER_SIZE)

        @stream.read(size).bytes.each_with_index do |byte, index|
          decrypted_stream << Unxls::Offcrypto._xor_decrypt_byte(byte, index, xor_array, pos, size)
        end
      end

    # https://github.com/nolze/msoffcrypto-tool/blob/master/msoffcrypto/format/xls97.py
    when :RC4
      salt = filepass_data[:Salt]
      encr_verifier = filepass_data[:EncryptedVerifier]
      encr_verifier_hash = filepass_data[:EncryptedVerifierHash]
      block_num = 0

      key = Unxls::Offcrypto._rc4_make_key(password, salt, block_num)
      ok = Unxls::Offcrypto._rc4_password_match?(key, encr_verifier, encr_verifier_hash)
      raise("Password '#{password}' does not match") unless ok

      @stream.rewind
      until @stream.eof?
        block_data = @stream.read(Unxls::Offcrypto::BLOCK_SIZE)
        key = Unxls::Offcrypto._rc4_make_key(password, salt, block_num)
        decrypted_stream << Unxls::Offcrypto._rc4_decrypt(block_data, key)
        block_num += 1
      end

    when :CryptoAPI
      salt = filepass_data[:EncryptionVerifier][:Salt]
      encr_verifier = filepass_data[:EncryptionVerifier][:EncryptedVerifier]
      encr_verifier_hash = filepass_data[:EncryptionVerifier][:EncryptedVerifierHash]
      hash_size = filepass_data[:EncryptionVerifier][:VerifierHashSize]
      key_size = filepass_data[:EncryptionHeader][:KeySize] / 8 # bits to bytes
      hash_alg = filepass_data[:_hashing_algorithm]
      unless (encr_alg = filepass_data[:_encryption_algorithm]) == :RC4
        # Couldn't find any examples of .xls files encrypted with AES cipher yet
        raise("This file's encryption type is not yet fully supported. Please contact the gem developer.")
      end
      block_num = 0

      key = Unxls::Offcrypto._rc4cryptoapi_make_key(password, salt, block_num, key_size, hash_alg)
      ok = Unxls::Offcrypto._rc4cryptoapi_password_match?(key, encr_verifier, encr_verifier_hash, hash_size, hash_alg, encr_alg)
      raise("Password '#{password}' does not match") unless ok

      @stream.rewind
      until @stream.eof?
        block_data = @stream.read(Unxls::Offcrypto::BLOCK_SIZE)
        key = Unxls::Offcrypto._rc4cryptoapi_make_key(password, salt, block_num, key_size, hash_alg)
        decrypted_stream << Unxls::Offcrypto._rc4cryptoapi_decrypt(block_data, key, encr_alg)
        block_num += 1
      end

    else
      raise "Decryption for type #{decryption_type} is not yet implemented"

    end

    @stream.close
    @stream = copy_unencrypted_parts(decrypted_stream)
  end

  # Overwrite the parts from the original stream that were originally unencrypted
  # see [MS-XLS].pdf, page 165
  # @param decrypted_stream [StringIO]
  # @return [StringIO]
  def copy_unencrypted_parts(decrypted_stream)
    @stream.rewind

    until @stream.eof?
      pos = @stream.pos
      id, size = read_record_header
      @stream.seek(pos)
      header = @stream.read(Unxls::Biff8::Record::HEADER_SIZE)
      data = @stream.read(size)

      decrypted_stream.seek(pos)
      case Unxls::Biff8::Record.name_by_id(id)
      when :BOF, :FilePass, :UsrExcl, :FileLock, :InterfaceHdr, :RRDInfo, :RRDHead
        decrypted_stream.write(header << data)
      when :BoundSheet8
        lb_ply_pos = data[0..3]
        decrypted_stream.write(header << lb_ply_pos)
      else
        decrypted_stream.write(header)
      end
    end

    decrypted_stream.rewind
    decrypted_stream
  end

  # @return [Unxls::Biff8::Record, nil]
  def get_next_record
    return if @stream.eof?

    pos = @stream.pos
    id, size = read_record_header

    return if id.zero? && last_parsed[:EOF] # skip zero padding of encrypted streams

    record_params = {
      id: id,
      pos: pos,
      size: size,
      data: [@stream.read(size)]
    }

    while record_continued?
      pos = @stream.pos
      cr_id, size = read_record_header

      record_params[:continue] ||= []
      record_params[:continue] << {
        id: cr_id,
        pos: pos,
        size: size
      }

      record_params[:data] << @stream.read(size)
    end

    Unxls::Log.debug_raw_record(record_params) # @debug

    Unxls::Biff8::Record.new(record_params, self)
  end

  def record_continued?
    return if @stream.eof?

    start_position = @stream.pos
    id, _ = read_record_header
    @stream.seek(start_position)

    Unxls::Biff8::Record.continue?(id)
  end

  # @return [Array<Integer>]
  def read_record_header
    @stream.read(Unxls::Biff8::Record::HEADER_SIZE).unpack('vv')
  end

  # @return [Hash]
  def last_parsed
    @parsed_result.last
  end

end
