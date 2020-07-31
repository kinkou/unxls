# frozen_string_literal: true

class Unxls::Parser

  attr_reader :file
  attr_accessor :settings

  # @param file [File]
  # @param settings [Hash]
  def initialize(file, settings)
    @file = file
    @settings = settings
  end

  # @return [Hash]
  def parse
    biff_version = detect_biff_version
    result = {}

    case biff_version
    when :BIFF8
      result[:workbook_stream] = Unxls::Biff8::WorkbookStream.new(self).parse
      # parse other storages/streams
    else
      raise "Sorry, #{biff_version} is not supported yet"
    end

    result
  end

  private

  # See "2.3 File Structure" of OpenOffice's doc, p.14
  # See "5.8 BOF – Beginning of File" of OpenOffice's doc, p.135
  # Test files: https://www.openoffice.org/sc/testdocs/
  # @return [Symbol]
  def detect_biff_version
    stream = begin
      compound = Ole::Storage.open(@file)
      Unxls::Log.debug(compound.dir.entries('.'), 'OLE compound file entries:', :cyan) # @debug

      if compound.file.exists?('Workbook') # BIFF8
        compound.file.open('Workbook')
      elsif compound.file.exists?('Book') # BIFF5
        compound.file.open('Book')
      else
        raise 'Error opening workbook stream'
      end
    rescue Ole::Storage::FormatError
      @file # Worksheet stream, BIFF2-4
    end

    id, _, vers, dt = stream.tap(&:rewind).read(8).unpack('v4')

    id_d = {
      0x0009 => :BIFF2,
      0x0209 => :BIFF3,
      0x0409 => :BIFF4,
      0x0809 => :BIFF5_8
    }[id]

    vers_d = {
      0x0000 => :BIFF5, # see 5.8.2 of OpenOffice's doc, p.136
      0x0200 => :BIFF2,
      0x0300 => :BIFF3,
      0x0400 => :BIFF4,
      0x0500 => :BIFF5,
      0x0600 => :BIFF8,
    }[vers]

    dt_d = {
      0x0005 => :globals, # Workbook globals substream
      0x0006 => :vb_module, # Visual Basic module substream
      0x0010 => :dialog_or_work_sheet, # If fDialog flag in the WsBool record in the substream is 1, it's a dialog sheet substream
      0x0020 => :chart, # Cart sheet substream
      0x0040 => :macro, # Macro sheet substream
      0x0100 => :workspace # Workspace substream
    }[dt]

    version_params = {
      stream_class: stream.class.to_s,
      id: Unxls::Log.h2b(id),
      id_d: id_d,
      vers: Unxls::Log.h2b(vers),
      vers_d: vers_d,
      dt: Unxls::Log.h2b(dt),
      dt_d: dt_d,
    }
    Unxls::Log.debug(version_params, 'Detecting BIFF version:', :cyan) # @debug

    if stream.is_a?(File) && %i(BIFF2 BIFF3 BIFF4).include?(id_d)
      id_d
    elsif id_d == :BIFF5_8 && vers_d && dt_d == :globals
      vers_d
    else
      raise 'Cannot detect BIFF version'
    end
  end

end