# frozen_string_literal: true

# [MS-OFFCRYPTO]: Office Document Cryptography Structure
module Unxls::Offcrypto
  extend self

  BLOCK_SIZE = 1024

  DEFAULT_PASSWORD = 'VelvetSweatshop'

  #
  # XOR obfuscation
  #

  XOR_PAD = [0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00]

  XOR_INITIAL_CODE = [
    0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C,
    0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139,
    0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3
  ]

  XOR_MATRIX = [
    0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09,
    0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF,
    0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0,
    0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40,
    0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5,
    0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A,
    0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9,
    0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0,
    0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC,
    0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10,
    0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168,
    0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C,
    0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD,
    0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC,
    0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4
  ]

  # 2.3.7.2 Binary Document XOR Array Initialization Method 1
  # FUNCTION CreateXorKey_Method1
  # @param password [String]
  # @return [Integer]
  def _create_xor_key_method1(password)
    ansi_password = _unicode_password_to_ansi(password)

    xor_key = XOR_INITIAL_CODE[ansi_password.size - 1]
    current_element = 0x68

    ansi_password.bytes.reverse.each do |c|
      7.times do
        xor_key ^= XOR_MATRIX[current_element] unless (c & 0x40).zero?
        c <<= 1
        current_element -= 1
      end
    end

    xor_key
  end

  # 2.3.7.2 Binary Document XOR Array Initialization Method 1
  # FUNCTION CreateXorArray_Method1
  # @param password [String]
  # @return [Array<Integer>]
  def _create_xor_array_method1(password)
    ansi_password = _unicode_password_to_ansi(password).bytes
    xor_key = _create_xor_key_method1(password)
    index = password.size
    obfuscation_array = Array.new(16) { 0 }

    xor_key_high = xor_key >> 8
    xor_key_low = xor_key & 0xFF

    if (index & 1) == 1
      obfuscation_array[index] = _xor_ror(XOR_PAD[0], xor_key_high)
      index -= 1
      obfuscation_array[index] = _xor_ror(ansi_password[-1], xor_key_low)
    end

    while index > 0
      index -= 1
      obfuscation_array[index] = _xor_ror(ansi_password[index], xor_key_high)
      index -= 1
      obfuscation_array[index] = _xor_ror(ansi_password[index], xor_key_low)
    end

    index = 15
    pad_index = 15 - password.size
    while pad_index > 0
      obfuscation_array[index] = _xor_ror(XOR_PAD[pad_index], xor_key_high)
      index -= 1
      pad_index -= 1
      obfuscation_array[index] = _xor_ror(XOR_PAD[pad_index], xor_key_low)
      index -= 1
      pad_index -= 1
    end

    obfuscation_array
  end

  # 2.3.7.1 Binary Document Password Verifier Derivation Method 1
  # FUNCTION CreatePasswordVerifier_Method1
  # @param password [String]
  # @return [Integer]
  def _xor_password_verifier_method1(password)
    verifier = 0
    ansi_password = _unicode_password_to_ansi(password)

    ansi_password.bytes.unshift(ansi_password.size).reverse.each do |b|
      int1 = (verifier & 0b0100_0000_0000_0000).zero? ? 0 : 1
      int2 = (verifier << 1) & 0b0111_1111_1111_1111
      int3 = int1 | int2
      verifier = int3 ^ b
    end

    verifier ^ 0xCE4B
  end

  # @param password [String]
  # @param verification_bytes [String]
  def _xor_password_match?(password, verification_bytes)
    _xor_password_verifier_method1(password) == verification_bytes
  end

  # See 2.3.7.4 Binary Document Password Verifier Derivation Method 2 (p. 61)
  # @note This method works reliably only for passwords consisting of ASCII characters
  # @param password [String]
  # @return [String]
  def _unicode_password_to_ansi(password)
    password.encode(Encoding::UTF_16LE).unpack('v*').map do |c|
      low_byte = c & 0xFF
      low_byte.zero? ? ((c >> 8) & 0xFF) : low_byte
    end.pack('C*')
  end

  def _xor_ror(byte1, byte2)
    Unxls::BitOps.new(byte1 ^ byte2).ror(8, 1)
  end

  # https://github.com/apache/poi/blob/trunk/src/java/org/apache/poi/poifs/crypt/CryptoFunctions.java#L445
  # https://github.com/apache/poi/blob/trunk/src/java/org/apache/poi/poifs/crypt/xor/XORDecryptor.java
  # @param byte [Integer]
  # @param byte_index [Integer]
  # @param xor_decryption_array [Array<Integer>]
  # @param record_offset [Integer]
  # @param record_size [Integer]
  # @return [String]
  def _xor_decrypt_byte(byte, byte_index, xor_decryption_array, record_offset, record_size)
    # "The initial value for XorArrayIndex is as follows:
    # XorArrayIndex = (FileOffset + Data.Length) % 16
    # The FileOffset variable in this context is the stream offset into the Workbook stream at
    # the time we are about to write each of the bytes of the record data.
    # This (the value) is then incremented after each byte is written."
    # From: http://social.msdn.microsoft.com/Forums/en-US/3dadbed3-0e68-4f11-8b43-3a2328d9ebd5
    xor_array_index = (record_offset + Unxls::Biff8::Record::HEADER_SIZE + record_size + byte_index) & 0xF

    byte ^= xor_decryption_array[xor_array_index]
    byte = ((byte << 3) | (byte >> 5)) & 0xFF # rotate right 5 bits

    [byte].pack('C')
  end

  #
  # RC4
  #

  # 2.1.4 Version
  # The Version structure specifies the version of a product or feature. It contains a major and a minor version number.
  # @param data [String]
  # @return [Hash]
  def version(data)
    v_major, v_minor = data.unpack('vv')

    {
      vMajor: v_major, # vMajor (2 bytes): An unsigned integer that specifies the major version number.
      vMinor: v_minor # vMinor (2 bytes): An unsigned integer that specifies the minor version number.
    }
  end

  # 2.3.6.1 RC4 Encryption Header
  # @param io [StringIO]
  # @return [Hash]
  def rc4encryptionheader(io)
    {
      EncryptionVersionInfo: version(io.read(4)), # EncryptionVersionInfo (4 bytes): A Version structure (section 2.1.4), where Version.vMajor MUST be 0x0001 and Version.vMinor MUST be 0x0001.
      Salt: io.read(16), # Salt (16 bytes): A randomly generated array of bytes that specifies the salt value used during password hash generation.
      EncryptedVerifier: io.read(16), # EncryptedVerifier (16 bytes): An additional 16-byte verifier encrypted using a 40-bit RC4 cipher initialized as specified in section 2.3.6.2, with a block number of 0x00000000.
      EncryptedVerifierHash: io.read(16) # EncryptedVerifierHash (16 bytes): A 40-bit RC4 encrypted MD5 hash of the verifier used to generate the EncryptedVerifier field.
    }
  end

  # 2.3.6.2 Encryption Key Derivation
  # @param password [String]
  # @param salt [String]
  # @param block_num [Integer]
  # @return [String]
  def _rc4_make_key(password, salt, block_num)
    h0 = OpenSSL::Digest::MD5.digest(password.encode(Encoding::UTF_16LE))
    buffer = (h0[0..4] + salt) * 16
    h1 = OpenSSL::Digest::MD5.digest(buffer)
    hfin = "#{h1[0..4]}#{[block_num].pack('V')}"
    OpenSSL::Digest::MD5.digest(hfin)
  end

  # 2.3.6.4 Password Verification
  # @param key [String]
  # @param encr_verifier [String]
  # @param encr_verifier_hash [String]
  def _rc4_password_match?(key, encr_verifier, encr_verifier_hash)
    cipher = OpenSSL::Cipher::RC4.new
    cipher.decrypt
    cipher.key = key
    decrypted_verifier = cipher.update(encr_verifier) + cipher.final
    decrypted_verifier_hash = cipher.update(encr_verifier_hash) + cipher.final
    hashed_verifier = OpenSSL::Digest::MD5.digest(decrypted_verifier)
    hashed_verifier == decrypted_verifier_hash
  end

  # @param data [String]
  # @param key [String]
  # @return [String]
  def _rc4_decrypt(data, key)
    cipher = OpenSSL::Cipher::RC4.new
    cipher.decrypt
    cipher.key = key
    cipher.update(data) + cipher.final
  end

  #
  # CryptoAPI
  #

  # 2.3.5.1 RC4 CryptoAPI Encryption Header
  # @param io [StringIO]
  # @return [Hash]
  def rc4cryptoapiheader(io)
    result = {
      EncryptionVersionInfo: version(io.read(4)), # EncryptionVersionInfo (4 bytes): A Version structure (section 2.1.4) that specifies the encryption version used to create the document and the encryption version required to open the document.
      EncryptionHeaderFlags: encryptionheaderflags(io.read(4).unpack('V').first) # EncryptionHeader.Flags (4 bytes): A copy of the Flags stored in the EncryptionHeader structure (section 2.3.2) that is stored in this stream.
    }

    encryption_header_size = io.read(4).unpack('V').first
    result[:EncryptionHeaderSize] = encryption_header_size # EncryptionHeaderSize (4 bytes): An unsigned integer that specifies the size, in bytes, of the EncryptionHeader structure.
    result[:EncryptionHeader] = encryptionheader(io) # EncryptionHeader (variable): An EncryptionHeader structure (section 2.3.2) used to encrypt the structure.
    result[:EncryptionVerifier] = encryptionverifier(io) # EncryptionVerifier (variable): An EncryptionVerifier structure as specified in section 2.3.3 that is generated as specified in section 2.3.5.5.
    result[:_encryption_algorithm] = _rc4cryptoapi_encryption_alg(result[:EncryptionHeader])
    result[:_hashing_algorithm] = _rc4cryptoapi_hashing_alg(result[:EncryptionHeader])

    result
  end

  # 2.3.2 EncryptionHeader
  # The EncryptionHeader structure is used by ECMA-376 document encryption [ECMA-376] and Office binary document RC4 CryptoAPI encryption, as defined in section 2.3.5, to specify encryption properties for an encrypted stream.
  # @param io [StringIO]
  # @return [Hash]
  def encryptionheader(io)
    result = {
      Flags: encryptionheaderflags(io.read(4).unpack('V').first), # Flags (4 bytes): An EncryptionHeaderFlags structure, as specified in section 2.3.1, that specifies properties of the encryption algorithm used.
      SizeExtra: io.read(4).unpack('V').first, # SizeExtra (4 bytes): A field that is reserved and for which the value MUST be 0x00000000.
    }

    alg_id = io.read(4).unpack('l<').first # AlgID (4 bytes): A signed integer that specifies the encryption algorithm.
    result[:AlgID] = alg_id
    result[:AlgID_d] = {
      0x0000 => :'Determined by Flags',
      0x6801 => :RC4,
      0x660E => :AES128,
      0x660F => :AES192,
      0x6610 => :AES256
    }[alg_id]

    alg_id_hash, key_size, provider_type, _, _ = io.read(20).unpack('l<VV')
    result[:AlgIDHash] = alg_id_hash # AlgIDHash (4 bytes): A signed integer that specifies the hashing algorithm together with the Flags.fExternal bit.
    result[:KeySize] = key_size # KeySize (4 bytes): An unsigned integer that specifies the number of bits in the encryption key.
    result[:ProviderType] = provider_type # ProviderType (4 bytes): An implementation-specific value that corresponds to constants accepted by the specified CSP.
    # Reserved1 (4 bytes): A value that is undefined and MUST be ignored.
    # Reserved2 (4 bytes): A value that MUST be 0x00000000 and MUST be ignored.

    result[:CSPName] = Unxls::Oshared._db_zero_terminated(io) # CSPName (variable): A null-terminated Unicode string that specifies the CSP name.

    result
  end

  # 2.3.1 EncryptionHeaderFlags
  # The EncryptionHeaderFlags structure specifies properties of the encryption algorithm used.
  # @param value [Integer]
  # @return [Hash]
  def encryptionheaderflags(value)
    attrs = Unxls::BitOps.new(value)

    {
      # A – Reserved1 (1 bit): A value that MUST be 0 and MUST be ignored.
      # B – Reserved2 (1 bit): A value that MUST be 0 and MUST be ignored.
      fCryptoAPI: attrs.set_at?(2), # C – fCryptoAPI (1 bit): A flag that specifies whether CryptoAPI RC4 or ECMA-376 encryption [ECMA- 376] is used.
      fDocProps: attrs.set_at?(3), # D – fDocProps (1 bit): A value that MUST be 0 if document properties are encrypted. The encryption of document properties is specified in section 2.3.5.4.
      fExternal: attrs.set_at?(4), # E – fExternal (1 bit): A value that MUST be 1 if extensible encryption is used.
      fAES: attrs.set_at?(5), # F – fAES (1 bit): A value that MUST be 1 if the protected content is an ECMA-376 document [ECMA- 376]
      # Unused (26 bits): A value that is undefined and MUST be ignored.
    }
  end

  # 2.3.3 EncryptionVerifier
  # @param io [StringIO]
  # @return [Hash]
  def encryptionverifier(io)
    {
      SaltSize: io.read(4).unpack('V').first, # SaltSize (4 bytes): An unsigned integer that specifies the size of the Salt field.
      Salt: io.read(16), # Salt (16 bytes): An array of bytes that specifies the salt value used during password hash generation.
      EncryptedVerifier: io.read(16), # EncryptedVerifier (16 bytes): A value that MUST be the randomly generated Verifier value encrypted using the algorithm chosen by the implementation.
      VerifierHashSize: io.read(4).unpack('V').first, # VerifierHashSize (4 bytes): An unsigned integer that specifies the number of bytes needed to contain the hash of the data used to generate the EncryptedVerifier field.
      EncryptedVerifierHash: io.read # EncryptedVerifierHash (variable): An array of bytes that contains the encrypted form of the hash of the randomly generated Verifier value.
    }
  end

  # See AlgIDHash description, p.33
  # @param encryption_header [Hash]
  # @return [Symbol]
  def _rc4cryptoapi_hashing_alg(encryption_header)
    combination = [
      encryption_header[:AlgIDHash],
      encryption_header[:Flags][:fExternal]
    ]

    case combination
    when [0x0000, true] then :'Determined by the application'

    when [0x0000, false],
         [0x8004, false] then :SHA1

    else raise "Unknown Flags and AlgIDHash combination: #{combination}"
    end
  end

  # See AlgID description, p.33
  # @param encryption_header [Hash]
  # @return [Symbol]
  def _rc4cryptoapi_encryption_alg(encryption_header)
    combination = [
      encryption_header[:Flags][:fCryptoAPI],
      encryption_header[:Flags][:fAES],
      encryption_header[:Flags][:fExternal],
      encryption_header[:AlgID]
    ]

    case combination
    when [false, false, true,  0x0000] then :'Determined by the application'

    when [true,  false, false, 0x0000],
         [true,  false, false, 0x6801] then :RC4

    when [true,  true,  false, 0x0000],
         [true,  true,  false, 0x660E] then :'AES-128-CBC'

    when [true,  true,  false, 0x660F] then :'AES-192-CBC'

    when [true,  true,  false, 0x6610] then :'AES-256-CBC'

    else raise "Unknown Flags and AlgID combination: #{combination}"
    end
  end

  # 2.3.5.2 RC4 CryptoAPI Encryption Key Generation
  # @param password [String]
  # @param salt [String]
  # @param block_num [Integer]
  # @param key_size [Integer]
  # @param hash_alg [Symbol]
  # @return [String]
  def _rc4cryptoapi_make_key(password, salt, block_num, key_size, hash_alg)
    h0_digest = _openssl_obj(:Digest, hash_alg)
    h0_digest.update(salt)
    h0_digest.update(password.encode(Encoding::UTF_16LE))
    h0 = h0_digest.digest

    h_digest = _openssl_obj(:Digest, hash_alg)
    h_digest.update(h0)
    h_digest.update([block_num].pack('V'))
    h = h_digest.digest

    h_fin = h[0..(key_size - 1)]
    h_fin << ("\x00" * (16 - key_size)) if key_size < 16

    h_fin
  end

  # 2.3.5.6 Password Verification
  # @param key [String]
  # @param encr_verifier [String]
  # @param encr_verifier_hash [String]
  # @param verifier_hash_size [Integer]
  # @param hash_alg [Symbol]
  # @param encr_alg [Symbol]
  def _rc4cryptoapi_password_match?(key, encr_verifier, encr_verifier_hash, verifier_hash_size, hash_alg, encr_alg)
    cipher = _openssl_obj(:Cipher, encr_alg)
    cipher.decrypt
    cipher.key = key
    decrypted_verifier = cipher.update(encr_verifier) + cipher.final
    decrypted_verifier_hash = cipher.update(encr_verifier_hash) + cipher.final
    decrypted_verifier_hash = decrypted_verifier_hash[0..(verifier_hash_size - 1)]
    hashed_verifier = _openssl_obj(:Digest, hash_alg).digest(decrypted_verifier)
    hashed_verifier == decrypted_verifier_hash
  end

  # @param data [String]
  # @param key [String]
  # @param encr_alg [Symbol]
  # @return [String]
  def _rc4cryptoapi_decrypt(data, key, encr_alg)
    cipher = _openssl_obj(:Cipher, encr_alg)
    cipher.decrypt
    cipher.key = key
    cipher.update(data) + cipher.final
  end

  # @param type [Symbol] e.g. :Cipher, :Digest
  # @param name [Symbol] e.g. :RC4, SHA1
  # @return [OpenSSL::Cipher, OpenSSL::Digest] e.g. OpenSSL::Cipher::RC4 instance
  def _openssl_obj(type, name)
    Module.class_eval("OpenSSL::#{type}").new(name.to_s)
  end

end
