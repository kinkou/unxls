# frozen_string_literal: true

class Unxls::Biff8::Record
  using Unxls::Helpers

  HEADER_SIZE = 4

  attr_reader :bytes
  attr_reader :params, :stream # @todo remove

  # @param params [Hash]
  # @option params [Integer] :id
  # @option params [Integer] :pos
  # @option params [Integer] :size
  # @option params [Array<String>] :data
  # @option params [Array<Hash>] :continue
  # @param stream [Unxls::Biff8::WorkbookStream]
  def initialize(params, stream)
    @params = params
    @stream = stream
    @serial = false # set to true for serial records like Font, Format etc.
    open_next_record_block
  end

  # @return [StringIO]
  def open_next_record_block
    @bytes = StringIO.new(@params[:data].shift || '')
  end

  # @return [Symbol, nil]
  def name
    self.class.name_by_id(@params[:id])
  end

  def end_of_data?
    @params[:data].size.zero? && @bytes.eof?
  end

  # @param id [Integer]
  # @return [Symbol]
  def self.name_by_id(id)
    Unxls::Biff8::Constants::RECORD_IDS[id]
  end

  # @return [Hash, nil]
  def process
    unless name
      warn("Unknown record with id #{Unxls::Log.h2b(@params[:id])}, ignoring")
      return
    end

    method_name = "r_#{name.downcase}".to_sym
    result = respond_to?(method_name) ? self.send(method_name) : nil
    return unless result

    result[:_record] = header_data
    Unxls::Log.debug(result, 'Parsed result:', :green) # @debug

    result
  end

  # @return [Hash]
  def header_data
    result = {
      id: @params[:id],
      name: name,
      pos: @params[:pos],
      size: @params[:size],
    }

    result[:index] = collection_index if serial?

    result
  end

  # See 2.1.4 Record, p. 57
  # @param id [Integer]
  def self.continue?(id)
    Unxls::Biff8::Constants::CONTINUE_RECORDS.include?(name_by_id(id))
  end

  def serial?
    @serial
  end

  # @return [Integer]
  def collection_index
    @stream.last_parsed[name] ? @stream.last_parsed[name].size : 0
  end

  # Record-processing methods
  #
  # +method naming+
  # processing methods for records from the 2.4 list: r_<record name downcase>, i.e. 'r_bof'
  #
  # +property naming+
  # orignal values named follow the specification, i.e. 'verXLHigh'
  # decoded (when meaning is unclear from the value itself) values named as <original value name>_d, i.e. 'verXLHigh_d'

  #
  # Globals substream records
  #

  # 2.4.21 BOF, page 212
  # The BOF record specifies the beginning of the individual substreams as specified by the workbook section. It also specifies history information for the substreams.
  # @return [Hash]
  def r_bof
    vers, dt, rup_build, rup_year = @bytes.read(8).unpack('v4')
    result = {
      vers: vers, # vers (2 bytes): An unsigned integer that specifies the BIFF version of the file. The value MUST be 0x0600.
      dt: dt, # dt (2 bytes): An unsigned integer that specifies the document type of the substream of records following this record.
      rupBuild: rup_build, # rupBuild (2 bytes): An unsigned integer that specifies the build identifier.
      rupYear: rup_year # rupYear (2 bytes): An unsigned integer that specifies the year when this BIFF version was first created. The value MUST be 0x07CC or 0x07CD.
    }

    result[:vers_d] = {
      0x0000 => :BIFF5, # see 5.8.2 of OpenOffice's doc, p.136
      0x0200 => :BIFF2,
      0x0300 => :BIFF3,
      0x0400 => :BIFF4,
      0x0500 => :BIFF5,
      0x0600 => :BIFF8,
    }[vers]

    result[:dt_d] = {
      0x0005 => :globals, # Specifies the workbook substream
      0x0006 => :vb_module, # Visual Basic module substream
      0x0010 => :dialog_or_work_sheet, # Specifies the dialog sheet substream or the worksheet substream. If fDialog flag in the WsBool record in the substream is 1, it's a dialog sheet substream
      0x0020 => :chart, # Cart sheet substream
      0x0040 => :macro, # Macro sheet substream
      0x0100 => :workspace # Workspace substream
    }[dt]

    attrs = Unxls::BitOps.new(@bytes.read(4).unpack('V').first)
    result[:fWin] = attrs.set_at?(0) # A - fWin (1 bit): A bit that specifies whether this file was last edited on a Windows platform. The value MUST be 1.
    result[:fRisc] = attrs.set_at?(1) # B - fRisc (1 bit): A bit that specifies whether the file was last edited on a RISC platform. The value MUST be 0.
    result[:fBeta] = attrs.set_at?(2) # C - fBeta (1 bit): A bit that specifies whether this file was last edited by a beta version of the application. The value MUST be 0.
    result[:fWinAny] = attrs.set_at?(3) # D - fWinAny (1 bit): A bit that specifies whether this file has ever been edited on a Windows platform. The value SHOULD<28> be 1.
    result[:fMacAny] = attrs.set_at?(4) # E - fMacAny (1 bit): A bit that specifies whether this file has ever been edited on a Macintosh platform. The value MUST be 0.
    result[:fBetaAny] = attrs.set_at?(5) # F - fBetaAny (1 bit): A bit that specifies whether this file has ever been edited by a beta version of the application. The value MUST be 0.
    # 6…7 G - unused1 (2 bits): Undefined and MUST be ignored.
    result[:fRiscAny] = attrs.set_at?(8) # H - fRiscAny (1 bit): A bit that specifies whether this file has ever been edited on a RISC platform. The value MUST be 0.
    result[:fOOM] = attrs.set_at?(9) # I - fOOM (1 bit): A bit that specifies whether this file had an out-of-memory failure.
    result[:fGlJmp] = attrs.set_at?(10) # J - fGlJmp (1 bit): A bit that specifies whether this file had an out-of-memory failure during rendering.
    # 11…12 K - unused2 (2 bits): Undefined, and MUST be ignored.
    result[:fFontLimit] = attrs.set_at?(13) # L - fFontLimit (1 bit): A bit that specified that whether this file hit the 255 font limit (this happens only for Excel 97)
    result[:verXLHigh] = attrs.value_at(14..17) # M - verXLHigh (4 bits): An unsigned integer that specifies the highest version of the application that once saved this file.
    result[:verXLHigh_d] = Unxls::Biff8::Structure._verxlhigh(result[:verXLHigh])
    # 18 N - unused3 (1 bit): Undefined, and MUST be ignored.
    # 19…31 reserved1 (13 bits): MUST be zero, and MUST be ignored.

    attrs_raw = @bytes.read(4).unpack('V').first
    attrs = Unxls::BitOps.new(attrs_raw)
    result[:verLowestBiff] = attrs.value_at(0..7) # verLowestBiff (8 bits): An unsigned integer that specifies the BIFF version saved. The value MUST be 6.
    result[:verLastXLSaved] = attrs.value_at(8..11) # O - verLastXLSaved (4 bits): An unsigned integer that specifies the application that saved this file most recently. The value MUST be the value of field verXLHigh or less.
    result[:verLastXLSaved_d] = Unxls::Biff8::Structure._verlastxlsaved(result[:verLastXLSaved])
    # 12…31 reserved2 (20 bits): MUST be zero, and MUST be ignored.

    result
  end

  # 2.4.28 BoundSheet8, page 220
  # The BoundSheet8 record specifies basic information about a sheet, including the sheet name, hidden state, and type of sheet.
  # @return [Hash]
  def r_boundsheet8
    @serial = true

    lb_ply_pos, attrs, dt = @bytes.read(6).unpack('VCC')
    hs_state = Unxls::BitOps.new(attrs).value_at(0..1)

    visibility = {
      0 => :visible,
      1 => :hidden,
      2 => :very_hidden # The sheet is hidden and cannot be displayed using the user interface.
    }[hs_state]

    # This is very similar to BOF.dt values table but still different
    type = {
      0 => :dialog_or_work_sheet, # The sheet substream that starts with the BOF record specified in lbPlyPos MUST contain one WsBool record. If the fDialog field in that WsBool is 1 then the sheet is dialog sheet. Otherwise, the sheet is a worksheet.
      1 => :macro,
      2 => :chart,
      6 => :vb_module
    }[dt]

    {
      lbPlyPos: lb_ply_pos, # lbPlyPos (4 bytes): A FilePointer as specified in [MS-OSHARED] section 2.2.1.5 that specifies the stream position of the start of the BOF record for the sheet.
      hsState: hs_state, # A - hsState (2 bits): An unsigned integer that specifies the hidden state of the sheet.
      hsState_d: visibility,
      # unused (6 bits): Undefined and MUST be ignored.
      dt: dt, # dt (8 bits): An unsigned integer that specifies the sheet type.
      dt_d: type,
      stName: Unxls::Biff8::Structure.shortxlunicodestring(@bytes) # stName (variable): A ShortXLUnicodeString structure that specifies the unique case-insensitive name of the sheet.
    }
  end

  # 2.4.35 CalcPrecision, page 224
  # The CalcPrecision record specifies the calculation precision mode for the workbook.
  # @return [Hash]
  def r_calcprecision
    {
      # If the value is 0, the precision as displayed mode is selected.
      # If the value is 1, the precision as displayed mode is not selected.
      fFullPrec: @bytes.read.unpack('v').first == 0 # fFullPrec (2 bytes): A Boolean (section 2.5.14) that specifies whether the precision as displayed mode is selected.
    }
  end

  # 2.4.52 CodePage, page 239
  # The CodePage record specifies code page information for the workbook.
  # @return [Hash]
  def r_codepage
    cv_data = @bytes.read.unpack('v').first

    {
      cv: cv_data, # cv (2 bytes): An unsigned integer that specifies the workbook’s code page. The value MUST be one of the code page values specified in [CODEPG] or the special value 1200, which means that the workbook is Unicode.
      cv_d: Unxls::Biff8::Constants::CODEPAGES[cv_data]
    }
  end

  # 2.4.63 Country, page 245
  # The Country record specifies locale information for a workbook.
  # @return [Hash]
  def r_country
    i_country_def, i_country_win_ini = @bytes.read(4).unpack('vv')

    {
      iCountryDef: i_country_def, # iCountryDef (2 bytes): An unsigned integer that specifies the country/region code determined by the locale in effect when the workbook was saved.
      iCountryDef_d: Unxls::Biff8::Constants::COUNTRIES[i_country_def],
      iCountryWinIni: i_country_win_ini, # iCountryWinIni (2 bytes): An unsigned integer that specifies the system regional settings country/region code in effect when the workbook was saved.
      iCountryWinIni_d: Unxls::Biff8::Constants::COUNTRIES[i_country_win_ini]
    }
  end

  # 2.4.77 Date1904, page 257
  # The Date1904 record specifies the date system that the workbook uses.
  # @return [Hash]
  def r_date1904
    {
      # 0: The workbook uses the 1900 date system. The first date of the 1900 date system is 00:00:00 on January 1, 1900, specified by a serial value of 1.
      # 1: The workbook uses the 1904 date system. The first date of the 1904 date system is 00:00:00 on January 1, 1904, specified by a serial value of 0.
      f1904DateSystem: @bytes.read.unpack('v').first == 1 # f1904DateSystem (2 bytes): A Boolean (section 2.5.14) that specifies the date system used in this workbook.
    }
  end

  # 2.4.97 DXF
  # The DXF record specifies a differential format.
  # @return [Hash]
  def r_dxf
    @serial = true

    frth_data = @bytes.read(12)
    attrs = @bytes.read(2).unpack('v').first
    attrs = Unxls::BitOps.new(attrs)

    result = {
      frtHeader: Unxls::Biff8::Structure.frtheader(frth_data), # frtHeader (12 bytes): An FrtHeader structure. The frtHeader.rt field MUST be 2189.
      # A - unused1 (1 bit): Undefined and MUST be ignored.
      fNewBorder: attrs.set_at?(1), # B - fNewBorder (1 bit): A bit that specifies whether it is possible to specify internal border formatting in xfprops. Internal border formatting is formatting that applies to borders that lie between a range of cells. Specifies that internal border formatting can be used in xfprops.
      # C - unused2 (1 bit): Undefined and MUST be ignored.
      # reserved (13 bits): MUST be zero, and MUST be ignored.
    }

    result.merge!(Unxls::Biff8::Structure.xfprops(@bytes)) # xfprops (variable): An XFProps structure that specifies the formatting properties.
  end

  # 2.4.117 FilePass
  # The FilePass record specifies the encryption algorithm used to encrypt the workbook and the structure that is used to verify the password provided when attempting to open the workbook.
  # @return [Hash]
  def r_filepass
    w_encryption_type = @bytes.read(2).unpack('v').first
    result = { wEncryptionType: w_encryption_type } # wEncryptionType (2 bytes): A Boolean (section 2.5.14) that specifies the encryption type.

    case w_encryption_type
    when 0x0000
      result[:_type] = :XOR
      result.merge!(Unxls::Biff8::Structure.xorobfuscation(@bytes)) # encryptionInfo (variable): A variable type field. The type and meaning of this field is dictated by the value of wEncryptionType.

    when 0x0001
      rc4_type = @bytes.read(2).unpack('v').first
      @bytes.pos -= 2

      case rc4_type
      when 0x0001
        result[:_type] = :RC4
        result.merge!(Unxls::Offcrypto.rc4encryptionheader(@bytes))

      when 0x0002, 0x0003, 0x0004
        result[:_type] = :CryptoAPI
        result.merge!(Unxls::Offcrypto.rc4cryptoapiheader(@bytes))

      else
        raise("Unknown RC4 encryption header type #{Unxls::Log.h2b(rc4_type)} in FilePass record")

      end

    else
      raise("Unknown encryption type #{Unxls::Log.h2b(w_encryption_type)} in FilePass record")

    end

    result
  end

  # 2.4.122 Font, page 298
  # The Font record specifies a font and font formatting information.
  # @return [Hash]
  def r_font
    @serial = true

    dy_height, attrs, icv, bls, sss, uls, b_family, b_char_set, _ = @bytes.read(14).unpack('v5C4')
    attrs = Unxls::BitOps.new(attrs)

    {
      dyHeight: dy_height, # dyHeight (2 bytes): An unsigned integer that specifies the height of the font in twips.
      # 0 A - unused1 (1 bit): Undefined and MUST be ignored.
      fItalic: attrs.set_at?(1), # B - fItalic (1 bit): A bit that specifies whether the font is italic.
      # 2 C - unused2 (1 bit): Undefined and MUST be ignored.
      fStrikeOut: attrs.set_at?(3), # D - fStrikeOut (1 bit): A bit that specifies whether the font has strikethrough formatting applied.

      # fOutline: attrs.set_at?(4), # (Seems like Mac Excel v.1,2 only) E - fOutline (1 bit): A bit that specifies whether the font has an outline effect applied.
      # fShadow: attrs.set_at?(5), # (Seems like Mac Excel v.1,2 only) F - fShadow (1 bit): A bit that specifies whether the font has a shadow effect applied.
      # fCondense: attrs.set_at?(6), # (Seems like there is no such option in Excel) G - fCondense (1 bit): A bit that specifies whether the font is condensed.
      # fExtend: attrs.set_at?(7), # (Seems like there is no such option in Excel) H - fExtend (1 bit): A bit that specifies whether the font is extended.

      # 8…15 reserved (8 bits): MUST be zero, and MUST be ignored.

      icv: icv, # icv (2 bytes): An unsigned integer that specifies the color of the font. The value SHOULD<88> be an IcvFont value.

      bls: bls, # bls (2 bytes): An unsigned integer that specifies the font weight.
      bls_d: Unxls::Biff8::Structure.bold(bls),

      sss: sss, # sss (2 bytes): An unsigned integer that specifies whether superscript, subscript, or normal script is used.
      sss_d: Unxls::Biff8::Structure.script(sss),

      uls: uls, # uls (1 byte): An unsigned integer that specifies the underline style.
      uls_d: Unxls::Biff8::Structure.underline(uls),

      bFamily: b_family, # bFamily (1 byte): An unsigned integer that specifies the font family this font belongs to.
      bFamily_d: Unxls::Biff8::Structure._font_family(b_family),

      bCharSet: b_char_set, # bCharSet (1 byte): An unsigned integer that specifies the character set.
      bCharSet_d: Unxls::Biff8::Structure._character_set(b_char_set),

      # unused3 (1 byte): Undefined and MUST be ignored.

      fontName: Unxls::Biff8::Structure.shortxlunicodestring(@bytes) # A ShortXLUnicodeString structure that specifies the name of this font.
    }
  end

  # 2.4.126 Format, page 302
  # The Format record specifies a number format.
  # See 18.8.31 numFmts (Number Formats) (p. 1774), Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf
  # @return [Hash]
  def r_format
    @serial = true

    ifmt = @bytes.read(2).unpack('v').first # See 2.5.165 IFmt

    {
      ifmt: ifmt, # ifmt (2 bytes): An IFmt structure that specifies the identifier of the format string specified by stFormat.
      stFormat: Unxls::Biff8::Structure.xlunicodestring(@bytes) # An XLUnicodeString structure that specifies the format string for this number format. The format string indicates how to format the numeric value of the cell.
    }
  end

  # 2.4.188 Palette, page 353
  # The Palette record specifies a custom color palette.
  # @return [Hash]
  def r_palette
    ccv = @bytes.read(2).unpack('s<').first

    result = {
      ccv: ccv, # ccv (2 bytes): A signed integer that specifies the number of colors in the rgColor array. The value MUST be 56.
      rgColor: [] # rgColor (variable): An array of LongRGB structures that specifies the colors of the color palette. The number of items in the array MUST be equal to the value specified in the ccv field.
    }

    ccv.times do
      result[:rgColor] << Unxls::Biff8::Structure.longrgba(@bytes.read(4))
    end

    result
  end

  # 2.4.265 SST, page 419
  # The SST record specifies string constants.
  # @return [Hash]
  def r_sst
    cst_total, cst_unique = @bytes.read(8).unpack('l<l<')
    result = {
      cstTotal: cst_total, # cstTotal (4 bytes): A signed integer that specifies the total number of references in the workbook to the strings in the shared string table.
      cstUnique: cst_unique, # cstUnique (4 bytes): A signed integer that specifies the number of unique strings in the shared string table.
      rgb: [] # rgb (variable): An array of XLUnicodeRichExtendedString structures. Records in this array are unique.
    }

    cst_unique.times do
      result[:rgb] << Unxls::Biff8::Structure.xlunicoderichextendedstring(self)
    end

    result
  end

  # 2.4.269 Style, page 426
  # The Style record specifies a cell style.
  # @return [Hash]
  def r_style
    @serial = true

    attrs = @bytes.read(2).unpack('v').first
    attrs = Unxls::BitOps.new(attrs)

    result = {
      ixfe: attrs.value_at(0..11), # ixfe (12 bits): An unsigned integer that specifies the zero-based index of the cell style XF in the collection of XF records in the Globals Substream.
      # A - unused (3 bits): Undefined and MUST be ignored.
      fBuiltIn: attrs.set_at?(15) # B - fBuiltIn (1 bit): A bit that specifies whether the cell style is built-in.
    }

    if result[:fBuiltIn]
      result[:builtInData] = Unxls::Biff8::Structure.builtinstyle(@bytes.read(2)) # builtInData (2 bytes): An optional BuiltInStyle structure that specifies the built-in cell style properties.
    else
      result[:user] = Unxls::Biff8::Structure.xlunicodestring(@bytes).to_sym # user (variable): An optional XLUnicodeString structure that specifies the name of the user-defined cell style.
    end

    result
  end

  # 2.4.270 StyleExt, page 427
  # The StyleExt record specifies additional information for a cell style.
  # @return [Hash]
  def r_styleext
    @serial = true

    result = { frtHeader: Unxls::Biff8::Structure.frtheader(@bytes.read(12)) }

    attrs, i_category = @bytes.read(2).unpack('CC')
    attrs = Unxls::BitOps.new(attrs)
    result[:fBuiltIn] = attrs.set_at?(0) # A - fBuiltIn (1 bit): A bit that specifies if this is a built-in cell style. If the value is 1, this is a built-in cell style. This value MUST match the fBuiltIn field of the preceding Style record.
    result[:fHidden] = attrs.set_at?(1) # B - fHidden (1 bit): A bit that specifies whether the cell style is not displayed in the user interface.
    result[:fCustom] = attrs.set_at?(2) # C - fCustom (1 bit): A bit that specifies whether the built-in cell style was modified by the user and thus has a custom definition.
    # reserved (5 bits): MUST be zero and MUST be ignored.
    result[:iCategory] = i_category # iCategory (1 byte): An unsigned integer that specifies which style category (2) that this style belongs to.
    result[:iCategory_d] = {
      0x00 => :'Custom style',
      0x01 => :'Good, bad, neutral style',
      0x02 => :'Data model style',
      0x03 => :'Title and heading style',
      0x04 => :'Themed cell style',
      0x05 => :'Number format style'
    }[i_category]

    built_in_data = @bytes.read(2)
    result[:builtInData] = Unxls::Biff8::Structure.builtinstyle(built_in_data) if result[:fBuiltIn] # builtInData (2 bytes): A BuiltInStyle structure that specifies the built-in cell style properties. If fBuiltIn is 0, this field MUST be 0xFFFF and MUST be ignored. If fBuiltIn is 1, this field MUST match the builtInData field of the preceding Style record.

    result[:stName] = Unxls::Biff8::Structure.lpwidestring(@bytes).to_sym # stName (variable): An LPWideString structure that specifies the name of the style to extend. MUST be less than or equal to 255 characters in length. If fBuiltIn is 0, the name specified by this field MUST match the name specified by the user field of the preceding Style record.

    result.merge!(Unxls::Biff8::Structure.xfprops(@bytes)) # xfProps (variable): An XFProps structure that specifies the formatting properties.
  end

  # 2.4.320 TableStyle
  # The TableStyle record specifies a user-defined table style and the beginning of a collection of TableStyleElement records as specified by the Globals Substream ABNF. The collection of TableStyleElement records specifies the properties of the table style.
  # @return [Hash]
  def r_tablestyle
    @serial = true

    frth_data = @bytes.read(12)
    attrs, ctse, cch_name = @bytes.read(8).unpack('vVv')
    attrs = Unxls::BitOps.new(attrs)

    {
      frtHeader: Unxls::Biff8::Structure.frtheader(frth_data), # frtHeader (12 bytes): An FrtHeader structure. The frtHeader.rt field MUST be 0x088F.
      # A - reserved1 (1 bit): MUST be zero, and MUST be ignored.
      fIsPivot: attrs.set_at?(1), # B - fIsPivot (1 bit): A bit that specifies whether the style can be applied to PivotTable views.
      fIsTable: attrs.set_at?(2), # C - fIsTable (1 bit): A bit that specifies whether the style can be applied to tables.
      # reserved2 (13 bits): MUST be zero, and MUST be ignored.
      ctse: ctse, # ctse (4 bytes): An unsigned integer that specifies the count of TableStyleElement records to follow this record. MUST be less than or equal to 28.
      cchName: cch_name, # cchName (2 bytes): An unsigned integer that specifies the count of characters in the rgchName field. This value MUST be less than or equal to 255 and greater than or equal to 1.
      rgchName: Unxls::Biff8::Structure._read_unicodestring(@bytes, cch_name, 1) # rgchName (variable): An array of Unicode characters whose length is specified by cchName that specifies the style name.
    }
  end

  # 2.4.321 TableStyleElement
  # The TableStyleElement record specifies formatting for one element of a table style. Each table style element specifies the formatting to apply to a particular area of a table or PivotTable view when the table style is applied.
  # @return [Hash]
  def r_tablestyleelement
    @serial = true

    frth_data = @bytes.read(12)
    tse_type, size, index = @bytes.read(12).unpack('V3')

    {
      frtHeader: Unxls::Biff8::Structure.frtheader(frth_data), # frtHeader (12 bytes): An FrtHeader structure. The frtHeader.rt field MUST be 0x0890.
      tseType: tse_type, # tseType (4 bytes): An unsigned integer that specifies the area of the table or PivotTable view to which the formatting is applied.
      tseType_d: Unxls::Biff8::Structure._tse_type(tse_type),
      size: size, # size (4 bytes): An unsigned integer that specifies the number of rows or columns to include in a single stripe band. MUST be ignored when the value of tseType does not equal 5, 6, 7, or 8. MUST be greater than or equal to 1 and less than or equal to 9.
      index: index, # index (4 bytes): A DXFId structure that specifies the DXF record that contains the differential formatting properties for this element.
      _tsi: @stream.last_parsed[:TableStyle].size - 1 # Index of parent TableStyle record
    }
  end

  # 2.4.322 TableStyles
  # The TableStyles record specifies the default table and PivotTable table styles and specifies the beginning of a collection of TableStyle records as defined by the Globals Substream ABNF. The collection of TableStyle records specifies user-defined table styles.
  # @return [Hash]
  def r_tablestyles
    frth_data = @bytes.read(12)
    cts, cch_table_style, cch_pivot_style = @bytes.read(8).unpack('Vvv')

    {
      frtHeader: Unxls::Biff8::Structure.frtheader(frth_data),
      cts: cts, # cts (4 bytes): An unsigned integer that specifies the total number of table styles in this document. This is the sum of the standard built-in table styles and all of the custom table styles.
      cchDefTableStyle: cch_table_style, # cchDefTableStyle (2 bytes): An unsigned integer that specifies the count of characters in the rgchDefTableStyle field.
      cchDefPivotStyle: cch_pivot_style, # cchDefPivotStyle (2 bytes): An unsigned integer that specifies the count of characters in the rgchDefPivotStyle field.
      rgchDefTableStyle: Unxls::Biff8::Structure._read_unicodestring(@bytes, cch_table_style, 1), # rgchDefTableStyle (variable): An array of Unicode characters whose length is specified by cchDefTableStyle that specifies the name of the default table style. Has double-byte characters
      rgchDefPivotStyle: Unxls::Biff8::Structure._read_unicodestring(@bytes, cch_pivot_style, 1), # rgchDefPivotStyle (variable): An array of Unicode characters whose length is specified by cchDefPivotStyle that specifies the name of the default PivotTable style. Has double-byte characters
    }
  end

  # 2.4.326 Theme, page 552
  # The Theme record specifies the theme in use in the document.
  # @return [Hash]
  def r_theme
    result = { frtHeader: Unxls::Biff8::Structure.frtheader(@bytes.read(12)) } # frtHeader (12 bytes): An FrtHeader structure. The value of the frtHeader.rt field MUST be 2198.

    dw_theme_version = @bytes.read(4).unpack('V').first
    result[:dwThemeVersion] = dw_theme_version # dwThemeVersion (4 bytes): An unsigned integer that specifies the theme type.
    result[:dwThemeVersion_d] = { 0 => :custom, 124226 => :default }[dw_theme_version]

    theme_data = @bytes.read
    while open_next_record_block.size > 0
      blocks.io.read(12) # skip frtHeader
      theme_data += blocks.io.read # rgb (variable): An optional byte stream that specifies the theme contents
    end

    # See 14.2.7 Theme Part (p. 135), Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf
    result[:rgb_d] = {} # unzipped theme xml files
    zip = Zip::InputStream.new(StringIO.new(theme_data))
    while (entry = zip.get_next_entry) do
      result[:rgb_d][entry.name] = entry.get_input_stream.read
    end

    result
  end

  # 2.4.353 XF, page 584
  # See also 2.5.282 XFIndex
  # The XF record specifies formatting properties for a cell or a cell style.
  # @return [Hash]
  def r_xf
    @serial = true

    ifnt, ifmt, attrs = @bytes.read(6).unpack('v3')
    attrs = Unxls::BitOps.new(attrs)
    result = {
      ifnt: ifnt, # ifnt (2 bytes): A FontIndex (2.5.129) structure that specifies a Font record.
      ifmt: ifmt, # ifmt (2 bytes): An IFmt (2.5.165) structure that specifies a number format identifier.
      fLocked: attrs.set_at?(0), # A - fLocked (1 bit): A bit that specifies whether the locked protection property is set to true.
      fHidden: attrs.set_at?(1), # B - fHidden (1 bit): A bit that specifies whether the hidden protection property is set to true.
      fStyle: attrs.set_at?(2), # C - fStyle (1 bit): A bit that specifies whether this record specifies a cell XF or a cell style XF. If the value is 1, this record specifies a cell style XF.
      f123Prefix: attrs.set_at?(3), # D - f123Prefix (1 bit): A bit that specifies whether prefix characters are present in the cell.
      ixfParent: attrs.value_at(4..15), # ixfParent (12 bits): An unsigned integer that specifies the zero-based index of a cell style XF record in the collection of XF records in the Globals Substream that this cell format inherits properties from. Cell style XF records are the subset of XF records with an fStyle field equal to 1. See XFIndex (2.5.282) for more information about the organization of XF records in the file. If fStyle equals 1, this field SHOULD equal 0xFFF, indicating there is no inheritance from a cell style XF.
    }

    specifies = Unxls::Biff8::Structure._builtin_xf_description(collection_index) # See 2.5.282 XFIndex
    result[:_description] = specifies if specifies

    result[:_type] = result[:fStyle] ? :stylexf : :cellxf

    result.merge(Unxls::Biff8::Structure.send(result[:_type], @bytes))
  end

  # 2.4.355 XFExt, page 585
  # The XFExt record specifies a set of formatting property extensions to an XF record in this file.
  # @return [Hash]
  def r_xfext
    @serial = true

    result = { frtHeader: Unxls::Biff8::Structure.frtheader(@bytes.read(12)) }
    result.merge(Unxls::Biff8::Structure._xfext(@bytes))
  end

  #
  # Worksheet substream records
  #

  # 2.4.20 Blank
  # The Blank record specifies an empty cell with no formula (section 2.2.2) or value.
  # @return [Hash]
  def r_blank
    @serial = true

    Unxls::Biff8::Structure.cell(@bytes.read) # cell (6 bytes): A Cell structure that specifies the cell.
  end

  # 2.4.24 BoolErr
  # The BoolErr record specifies a cell that contains either a Boolean value or an error value.
  # @return [Hash]
  def r_boolerr
    @serial = true

    {
      **Unxls::Biff8::Structure.cell(@bytes.read(6)), # cell (6 bytes): A Cell structure that specifies the cell.
      bes: Unxls::Biff8::Structure.bes(@bytes.read(2)) # bes (2 bytes): A Bes structure that specifies a Boolean or an error value.
    }
  end

  # 2.4.42 CF, page 228
  # The CF record specifies a conditional formatting rule.
  # @return [Hash]
  def r_cf
    @serial = true

    ct, cp, cce1, cce2 = @bytes.read(6).unpack('CCvv')

    ct_d = {
      0x01 => :cp_function_true, # Apply the conditional formatting when the comparison function specified by cp applied to the cell value, rgce1 and rgce2, evaluates to TRUE.
      0x02 => :rgce1_formula_true # Apply the conditional formatting when the formula (section 2.2.2) specified by rgce1 evaluates to TRUE.
    }[ct]

    {
      ct: ct, # An unsigned integer that specifies the type of condition
      ct_d: ct_d,
      cp: cp, # An unsigned integer that specifies the comparison function used when ct is equal to 0x01. In the following table, v represents the cell value, and v1 and v2 represent the results of evaluating the formulas specified by rgce1 and rgce2.
      cp_d: Unxls::Biff8::Structure._cf_cp(cp),
      cce1: cce1, # An unsigned integer that specifies the size of rgce1 in bytes.
      cce2: cce2, # An unsigned integer that specifies the size of rgce2 in bytes.
      rgbdxf: Unxls::Biff8::Structure.dxfn(@bytes), # A DXFN structure that specifies the formatting to apply to a cell that fulfills the condition.
      rgce1: Unxls::Biff8::Structure.cfparsedformulanocce(@bytes.read(cce1)), # A CFParsedFormulaNoCCE structure that specifies the first formula. If ct is equal to 0x01, this field is the first operand of the comparison. If ct is equal to 0x02, this formula is used to determine if the conditional formatting is applied.
      rgce2: Unxls::Biff8::Structure.cfparsedformulanocce(@bytes.read(cce2)) # A CFParsedFormulaNoCCE structure that specifies the formula that is the second operand of the comparison if ct is equal to 0x01 and cp is either equal to 0x01 or 0x02.
    }
  end

  # 2.4.43 CF12
  # The CF12 record specifies a conditional formatting rule.
  # @return [Hash]
  def r_cf12
    @serial = true

    result = {}
    result[:frtRefHeader] = Unxls::Biff8::Structure.frtrefheader(@bytes) # frtRefHeader (12 bytes): An FrtRefHeader.

    ct, cp, cce1, cce2 = @bytes.read(6).unpack('CCvv')
    ct_d = {
      0x01 => :cp_function_true, # Apply the conditional formatting if the comparison operation specified by cp evaluates to TRUE. rgbCT MUST be omitted.
      0x02 => :rgce1_formula_true, # Apply the conditional formatting if the formula (section 2.2.2) specified by rgce1 evaluates to TRUE. rgbCT MUST be omitted.
      0x03 => :color_scale_formatting, # Use color scale formatting. rgbCT is a CFGradient.
      0x04 => :data_bar_formatting, # Use data bar formatting. rgbCT is a CFDatabar.
      0x05 => :passes_cffilter, # Apply the conditional formatting when the cell value passes a filter specified in the rgbCT structure. rgbCT is a CFFilter.
      0x06 => :icon_set_formatting, # Use icon set formatting. rgbCT is a CFMultistate.
    }[ct]

    result.merge!({
      ct: ct, # ct (1 byte): An unsigned integer that specifies the type of condition. This field determines the type of the rgbCT field.
      ct_d: ct_d,
      cp: cp, # cp (1 byte): An unsigned integer that specifies the comparison function used when ct is equal to 0x01.
      cp_d: Unxls::Biff8::Structure._cf_cp(cp),
      cce1: cce1, # cce1 (2 bytes): An unsigned integer that specifies the size of rgce1 in bytes.
      cce2: cce2, # cce2 (2 bytes): An unsigned integer that specifies the size of rgce2 in bytes.
    })

    result[:dxf] = Unxls::Biff8::Structure.dxfn12(@bytes) # dxf (variable): A DXFN12 that specifies the formatting to apply to a cell that fulfills the condition.

    # Rest of the record not implemented yet:
    result[:rgce1] = Unxls::Biff8::Structure.cfparsedformulanocce(@bytes.read(cce1)) # rgce1 (variable): A CFParsedFormulaNoCCE that specifies the formula used to evaluate the first operand in a comparison when ct is 0x01. If ct is 0x02 rgce1 MUST be a Boolean function.
    result[:rgce2] = Unxls::Biff8::Structure.cfparsedformulanocce(@bytes.read(cce2)) # rgce2 (variable): A CFParsedFormulaNoCCE that specifies the formula used to evaluate the second operand of the comparison when ct is 0x01 and cp is either 0x01 or 0x02.
    result[:fmlaActive] = Unxls::Biff8::Structure.cfparsedformula(@bytes) # fmlaActive (variable): A CFParsedFormula that specifies the formula that specifies an activity condition for the color scale, data bar and icon set formatting rule types. If ct is equal to 0x03, 0x04 or 0x06, then the conditional formatting is applied if fmlaActive evaluates to TRUE.

    a_e_raw, ipriority, icf_template, cb_template_parm = @bytes.read(6).unpack('CvvC')
    a_e = Unxls::BitOps.new(a_e_raw)

    # 0 A - unused1 (1 bit): Undefined and MUST be ignored.
    result[:fStopIfTrue] = a_e.set_at?(1) # 1 B - fStopIfTrue (1 bit): A bit that specifies whether, when a cell fulfills the condition corresponding to this rule, the lower priority conditional formatting rules that apply to this cell are evaluated.
    # 2…3 C - reserved1 (2 bits): MUST be zero and MUST be ignored.
    # 4 D - unused2 (1 bit): Undefined and MUST be ignored.
    # 5…7 E - reserved2 (3 bits): MUST be zero and MUST be ignored.

    result[:ipriority] = ipriority # ipriority (2 bytes): An unsigned integer that specifies the priority of the rule. Rules that apply to the same cell are evaluated in increasing order of ipriority. MUST be unique across all CF12 records and CFExNonCF12 structures in the worksheet substream.

    result[:icfTemplate] = icf_template # icfTemplate (2 bytes): An unsigned integer that specifies the template from which the rule was created.
    result[:icfTemplate_d] = Unxls::Biff8::Structure._icf_template_d(icf_template)

    result[:cbTemplateParm] = cb_template_parm # cbTemplateParm (1 byte): An unsigned integer that specifies the size of the rgbTemplateParms field in bytes. MUST be 16.
    result[:rgbTemplateParms] = Unxls::Biff8::Structure.cfextemplateparams(@bytes.read(cb_template_parm), icf_template) # rgbTemplateParms (16 bytes): A CFExTemplateParams that specifies the parameters for the rule.

    # rgbCT (variable): A field that specifies the parameters of this rule. The type of rgbCT depends on the value of ct.
    result[:rgbCT] = case ct
    # Apply the conditional formatting if the comparison operation specified by cp evaluates to TRUE
    # Apply the conditional formatting if the formula (section 2.2.2) specified by rgce1 evaluates to TRUE.
    when 0x01, 0x02 then :omitted # rgbCT MUST be omitted.
    when 0x03 then Unxls::Biff8::Structure.cfgradient(@bytes) # Use color scale formatting.
    when 0x04 then Unxls::Biff8::Structure.cfdatabar(@bytes) # Use data bar formatting.
    when 0x05 then Unxls::Biff8::Structure.cffilter(@bytes) # Apply the conditional formatting when the cell value passes a filter specified in the rgbCT structure.
    when 0x06 then Unxls::Biff8::Structure.cfmultistate(@bytes) # Use icon set formatting.
    else raise "Unexpected value '#{ct}' of ct field in CF12 record"
    end

    result
  end

  # 2.4.44 CFEx
  # The CFEx record extends a CondFmt.
  # @return [Hash]
  def r_cfex
    @serial = true

    result = {
      frtRefHeaderU: Unxls::Biff8::Structure.frtrefheaderu(@bytes) # frtRefHeaderU (12 bytes): An FrtRefHeaderU structure.
    }

    f_is_cf12, n_id = @bytes.read(6).unpack('Vv')
    # fIsCF12 (4 bytes): A Boolean (section 2.5.14) that specifies what type of rule this record extends. MUST be one of the following values:
    # 0x00000000 – This record extends a rule specified by a CF record and MUST NOT be followed by a CF12 record.
    # 0x00000001 – This record extends a rule specified by a CF12 record and MUST be followed by the CF12 record it extends.
    result[:fIsCF12] = f_is_cf12
    result[:nID] = n_id # nID (2 bytes): An unsigned integer that specifies which CondFmt record is being extended. It MUST be equal to the nID field of one of the CondFmt records in the Worksheet substream.

    if f_is_cf12.zero?
      result[:rgbContent] = Unxls::Biff8::Structure.cfexnoncf12(@bytes) # rgbContent (variable): A CFExNonCF12 structure that specifies the extensions to an existing CF record. MUST be omitted when fIsCF12 is not equal to 0x00.
    end

    result
  end

  # 2.4.53 ColInfo, page 240
  # The ColInfo record specifies the column formatting for a range of columns.
  # @return [Hash]
  def r_colinfo
    @serial = true

    col_first, col_last, coldx, ixfe, attrs, _ = @bytes.read.unpack('v6')
    attrs = Unxls::BitOps.new(attrs)

    {
      colFirst: col_first, # colFirst (2 bytes): A Col256U structure that specifies the first formatted column.
      colLast: col_last, # colLast (2 bytes): A Col256U structure that specifies the last formatted column. The value MUST be greater than or equal to colFirst.
      coldx: coldx, # coldx (2 bytes): An unsigned integer that specifies the column width in units of 1/256th of a character width. Character width is defined as the maximum digit width of the numbers 0, 1, 2, … 9 as rendered in the Normal style’s font.
      ixfe: ixfe, # ixfe (2 bytes):  An IXFCell structure that specifies the default format for the column cells.
      fHidden: attrs.set_at?(0), # A - fHidden (1 bit): A bit that specifies whether the column range defined by colFirst and colLast is hidden.
      fUserSet: attrs.set_at?(1), # B - fUserSet (1 bit): A bit that specifies that the column width was either manually set by the user or is different from the default column width as specified by DefColWidth. If the value is 1, the column width was manually set or is different from DefColWidth.
      fBestFit: attrs.set_at?(2), # C - fBestFit (1 bit): A bit that specifies whether the column range defined by colFirst and colLast is set to "best fit." "Best fit" implies that the column width resizes based on the cell contents, and that the column width does not equal the default column width as specified by DefColWidth.
      fPhonetic: attrs.set_at?(3), # D - fPhonetic (1 bit): A bit that specifies whether phonetic information is displayed by default for the column range defined by colFirst and colLast.
      # 4…7 E - reserved1 (4 bits):  MUST be zero, and MUST be ignored.
      iOutLevel: attrs.value_at(8..10), # F - iOutLevel (3 bits): An unsigned integer that specifies the outline level of the column range defined by colFirst and colLast.
      # G - unused1 (1 bit):  Undefined and MUST be ignored.
      fCollapsed: attrs.set_at?(12) # H - fCollapsed (1 bit): A bit that specifies whether the column range defined by colFirst and colLast is in a collapsed outline state.
      # I - reserved2 (3 bits): MUST be zero, and MUST be ignored.
      # unused2 (2 bytes): Undefined and MUST be ignored.
    }
  end

  # 2.4.56 CondFmt, page 242
  # @return [Hash]
  def r_condfmt
    @serial = true

    ccf, attrs = @bytes.read(4).unpack('vv')
    attrs = Unxls::BitOps.new(attrs)

    {
      ccf: ccf, # An unsigned integer that specifies the count of CF records that follow this record.
      fToughRecalc: attrs.set_at?(0), # A bit that specifies that the appearance of the cell requires significant processing.
      nID: attrs.value_at(1..15), # An unsigned integer that identifies this record. The CFEx record uses this identifier to specify which CondFmt it extends.
      refBound: Unxls::Biff8::Structure.ref8u(@bytes.read(8)), # A Ref8U structure that specifies the bounds of the set of cells to which the conditional formatting rules apply.
      sqref: Unxls::Biff8::Structure.sqrefu(@bytes) # A SqRefU structure that specifies the cells to which the conditional formatting rules apply.
    }
  end

  # 2.4.57 CondFmt12, page 242
  # @return [Hash]
  def r_condfmt12
    @serial = true

    {
      frtRefHeaderU: Unxls::Biff8::Structure.frtrefheaderu(@bytes), # frtRefHeaderU (12 bytes): An FrtRefHeaderU structure
      mainCF: Unxls::Biff8::Structure.condfmtstructure(@bytes) # mainCF (variable): A CondFmtStructure structure that specifies properties of a set of conditional formatting rules.
    }
  end

  # 2.4.90 Dimensions
  # The Dimensions record specifies the used range of the sheet. It specifies the row and column bounds of used cells in the sheet. Used cells include all cells with formulas (section 2.2.2) or data. Used cells also include all cells with formatting applied directly to the cell. …
  # @return [Hash]
  def r_dimensions
    rw_mic, rw_mac, col_mic, col_mac, _ = @bytes.read.unpack('VVv3')

    {
      rwMic: rw_mic, # rwMic (4 bytes): A RwLongU structure that specifies the first row in the sheet that contains a used cell.
      rwMac: rw_mac, # rwMac (4 bytes): An unsigned integer that specifies the zero-based index of the row after the last row in the sheet that contains a used cell. MUST be less than or equal to 0x00010000. If this value is 0x00000000, no cells on the sheet are used cells.
      colMic: col_mic, # colMic (2 bytes): A ColU structure that specifies the first column in the sheet that contains a used cell.
      colMac: col_mac, # colMac (2 bytes): An unsigned integer that specifies the zero-based index of the column after the last column in the sheet that contains a used cell. MUST be less than or equal to 0x0100. If this value is 0x0000, no cells on the sheet are used cells.
      # reserved (2 bytes):  MUST be zero, and MUST be ignored.
    }
  end

  # @todo May be Continue-ed (!!!)
  # 2.4.114 Feature11
  # The Feature11 record specifies specific shared feature data. The only shared feature type stored in this record is a table in a worksheet.
  # @return [Hash]
  def r_feature11
    @serial = true

    frt_ref_header_u = Unxls::Biff8::Structure.frtrefheaderu(@bytes)
    isf, _, _, cref2, cb_feat_data, _ = @bytes.read(15).unpack('vCVvVv')

    {
      frtRefHeaderU: frt_ref_header_u, # frtRefHeaderU (12 bytes): An FrtRefHeaderU. The frtRefHeaderU.rt field MUST be 0x0872. The frtRefHeaderU.ref8 MUST refer to a range of cells associated with this record.
      isf: isf, # isf (2 bytes): A SharedFeatureType enumeration that specifies the type of Shared Feature data stored in the rgbFeat field. MUST be ISFLIST.
      isf_d: Unxls::Biff8::Structure.sharedfeaturetype(isf),
      # reserved1 (1 byte): Reserved and MUST be zero.
      # reserved2 (4 bytes): MUST be zero, and MUST be ignored.
      cref2: cref2, # cref2 (2 bytes): An unsigned integer that specifies the count of Ref8U records within the refs2 field.
      cbFeatData: cb_feat_data, # cbFeatData (4 bytes): An unsigned integer that specifies the size in bytes of the rgbFeat variable-size field. If the value is 0x0000, the size of the rgbFeat field is calculated by the following formula: size of rgbFeat = total size of record in bytes – size of refs2 in bytes – 27 bytes
      # reserved3 (2 bytes): MUST be zero, and MUST be ignored.
      refs2: cref2.times.map { Unxls::Biff8::Structure.ref8u(@bytes.read(8)) }, # refs2 (variable): An array of Ref8U structures that specifies references to ranges of cells within the worksheet associated with the feature. The count of records within this field is specified by the cref2 field.
      rgbFeat: Unxls::Biff8::Structure.tablefeaturetype(@bytes) # rgbFeat (variable): A variable-size structure that contains feature specific data. The size of the structure is specified by the cbFeatData field. This field MUST contain a TableFeatureType structure.
    }
  end

  # 2.4.115 Feature12
  # The Feature12 record specifies shared feature data that is used to describe a table in a worksheet. This record is used to encapsulate a table that has properties not supported by the Feature11 record.
  # @return [Hash]
  alias :r_feature12 :r_feature11

  # 2.4.127 Formula
  # The Formula record specifies a formula (section 2.2.2) for a cell.
  # @return [Hash]
  def r_formula
    @serial = true

    result = {
      **Unxls::Biff8::Structure.cell(@bytes.read(6)), # cell (6 bytes): A Cell structure that specifies a cell on the sheet.
      **Unxls::Biff8::Structure.formulavalue(@bytes.read(8)) # val (8 bytes): A FormulaValue structure that specifies the value of the formula.
    }

    attrs = Unxls::BitOps.new(@bytes.read(2).unpack('v').first)
    @bytes.read(4) # chn

    result.merge!({
      fAlwaysCalc: attrs.set_at?(0), # A - fAlwaysCalc (1 bit): A bit that specifies whether the formula needs to be calculated during the next recalculation.
      # B - reserved1 (1 bit): MUST be zero, and MUST be ignored.
      fFill: attrs.set_at?(2), # C - fFill (1 bit): A bit that specifies whether the cell has a fill alignment or a center-across-selection alignment.
      fShrFmla: attrs.set_at?(3), # D - fShrFmla (1 bit): A bit that specifies whether the formula is part of a shared formula as defined in ShrFmla. If this formula is part of a shared formula, formula.rgce MUST begin with a PtgExp structure.
      # E - reserved2 (1 bit): MUST be zero, and MUST be ignored.
      fClearErrors: attrs.set_at?(5), # F - fClearErrors (1 bit): A bit that specifies whether the formula is excluded from formula error checking.
      # reserved3 (10 bits): MUST be zero, and MUST be ignored.
      chn: :not_implemented, # chn (4 bytes): A field that specifies an application-specific cache of information. This cache exists for performance reasons only, and can be rebuilt based on information stored elsewhere in the file without affecting calculation results.
      formula: :not_implemented # formula (variable): A CellParsedFormula structure that specifies the formula.
    })
  end

  # 2.4.140 HLink
  # The HLink record specifies a hyperlink associated with a range of cells.
  # @return [Hash]
  def r_hlink
    @serial = true

    result = Unxls::Biff8::Structure.ref8u(@bytes.read(8)) # ref8 (8 bytes): A Ref8U structure that specifies the range of cells containing the hyperlink.
    result[:hlinkClsid] = Unxls::Dtyp.guid(@bytes.read(16)) # hlinkClsid (16 bytes): A class identifier (CLSID) (see RFC 4122, https://www.rfc-editor.org/rfc/rfc4122.txt) that specifies the COM component which saved the Hyperlink Object (as defined by [MS-OSHARED] section 2.3.7.1) in hyperlink.
    result[:hyperlink] = Unxls::Oshared.hyperlink(@bytes) # hyperlink (variable): A Hyperlink Object (as defined by [MS-OSHARED] section 2.3.7.1) that specifies the hyperlink and hyperlink-related information.

    result
  end

  # 2.4.141 HLinkTooltip
  # The HLinkTooltip record specifies the hyperlink ToolTip associated with a range of cells.
  # @return [Hash]
  def r_hlinktooltip
    @serial = true

    {
      frtRefHeaderNoGrbit: Unxls::Biff8::Structure.frtrefheadernogrbit(@bytes), # frtRefHeaderNoGrbit (10 bytes): An FrtRefHeaderNoGrbit structure. The frtRefHeaderNoGrbit.rt field MUST be 0x0800. The frtRefHeaderNoGrbit.ref8 field MUST match a Ref8U field from an existing HLink record.
      wzTooltip: Unxls::Oshared._db_zero_terminated(@bytes) # wzTooltip (variable): An array of Unicode characters that specifies the ToolTip string. String length MUST be greater than or equal to 2 and less than or equal to 256 (inclusive of null terminator) and the string MUST be null-terminated.
    }
  end

  # 2.4.149 LabelSst
  # The LabelSst record specifies a cell that contains a string.
  # @return [Hash]
  def r_labelsst
    @serial = true

    {
      **Unxls::Biff8::Structure.cell(@bytes.read(6)), # cell (6 bytes): A Cell structure that specifies the cell containing the string from the shared string table.
      isst: @bytes.read(4).unpack('V').first # isst (4 bytes): An unsigned integer that specifies the zero-based index of an element in the array of XLUnicodeRichExtendedString structure in the rgb field of the SST record in this Workbook Stream ABNF that specifies the string contained in the cell. MUST be greater than or equal to zero and less than the number of elements in the rgb field of the SST record.
    }
  end

  # 2.4.157 List12
  # The List12 record specifies the additional formatting information for a table. These records immediately follow a Feature11 or Feature12 record, and specify additional formatting information for the table specified by the Feature11 or Feature12 record. This record is a future record type record.
  # @return [Hash]
  def r_list12
    @serial = true

    frth_data = @bytes.read(12)
    lsd, id_list = @bytes.read(6).unpack('vV')

    lsd_d = {
      0x00 => :List12BlockLevel, # rgb is a List12BlockLevel structure that specifies the table block-level formatting.
      0x01 => :List12TableStyleClientInfo, # rgb is a List12TableStyleClientInfo structure that specifies the table style.
      0x02 => :List12DisplayName # rgb is a List12DisplayName structure that specifies the display name.
    }[lsd]

    method_name = lsd_d.to_s.downcase.to_sym

    result = {
      frtHeader: Unxls::Biff8::Structure.frtheader(frth_data), # frtHeader (12 bytes): An FrtHeader structure. The frtHeader.rt field MUST be 0x0890.
      lsd: lsd, # lsd (2 bytes): An unsigned integer that specifies the type of data contained in the rgb field. MUST be a value specified in the table listed under rgb.
      lsd_d: lsd_d,
      idList: id_list, # idList (4 bytes): An unsigned integer that identifies the associated table for which this record specifies additional formatting. MUST NOT be zero. MUST be equal to the idList field of the TableFeatureType structure embedded in the associated Feature11 or Feature12 record.
    }

    result[:rgb] = Unxls::Biff8::Structure.send(method_name, @bytes)

    result
  end

  # 2.4.168 MergeCells
  # The MergeCells record specifies merged cells in the document. If the count of the merged cells in the document is greater than 1026, the file will contain multiple adjacent MergeCells records.
  # @return [Hash]
  def r_mergecells
    @serial =  true

    result = {
      cmcs: @bytes.read(2).unpack('v').first, # cmcs (2 bytes): An unsigned integer that specifies the count of Ref8 structures. MUST be less than or equal to 1026.
      rgref: []
    }

    result[:cmcs].times do
      result[:rgref] << Unxls::Biff8::Structure.ref8(@bytes.read(8)) # rgref (variable): An array of Ref8 structures. Each array element specifies a range of cells that are merged into a single merged cell. These ranges MUST NOT overlap. MUST contain the number of elements specified by cmcs.
    end

    result
  end

  # 2.4.174 MulBlank
  # The MulBlank record specifies a series of blank cells in a sheet row. This record can store up to 256 IXFCell structures.
  # @return [Hash]
  def r_mulblank
    @serial = true

    rw, col_first = @bytes.read(4).unpack('vv')
    rgixfe_collast_data = @bytes.read # rest of the record
    col_last = rgixfe_collast_data[-2..-1].unpack('v').first # last 2 bytes
    rgixfe = rgixfe_collast_data[0..-3].unpack('v*') # all except last 2 bytes

    {
      rw: rw, # rw (2 bytes): An Rw structure that specifies a row containing the blank cells.
      colFirst: col_first, # colFirst (2 bytes): A Col structure that specifies the first column in the series of blank cells within the sheet. The value of colFirst.col MUST be less than or equal to 254.
      colLast: col_last, # colLast (2 bytes): A Col structure that specifies the last column in the series of blank cells within the sheet. This colLast.col value MUST be greater than colFirst.col value.
      rgixfe: rgixfe # rgixfe (variable): An array of IXFCell structures. Each element of this array contains an IXFCell structure corresponding to a blank cell in the series. The number of entries in the array MUST be equal to the value given by the following formula: Number of entries in rgixfe = (colLast.col – colFirst.col +1)
    }
  end

  # 2.4.175 MulRk
  # The MulRk record specifies a series of cells with numeric data in a sheet row. This record can store up to 256 RkRec structures.
  # @return [Hash]
  def r_mulrk
    @serial = true

    rw, col_first = @bytes.read(4).unpack('vv')
    rgrkrec_collast_data = @bytes.read # rest of the record
    col_last = rgrkrec_collast_data[-2..-1].unpack('v').first # last 2 bytes

    result = {
      rw: rw, # rw (2 bytes): An Rw structure that specifies the row containing the cells with numeric data.
      colFirst: col_first, # colFirst (2 bytes): A Col structure that specifies the first column in the series of numeric cells within the sheet. The value of colFirst.col MUST be less than or equal to 254.
      colLast: col_last, # colLast (2 bytes): A Col structure that specifies the last column in the set of numeric cells within the sheet. This colLast.col value MUST be greater than the colFirst.col value.
      rgrkrec: [] # rgrkrec (variable): An array of RkRec structures. Each element in the array specifies an RkRec in the row. The number of entries in the array MUST be equal to the value given by the following formula: Number of entries in rgrkrec = (colLast.col – colFirst.col +1)
    }

    rgrkrec_array_data_io = rgrkrec_collast_data[0..-3].to_sio # all except last 2 bytes
    while (rkrec = rgrkrec_array_data_io.read(6))
      result[:rgrkrec] << Unxls::Biff8::Structure.rkrec(rkrec)
    end

    result
  end

  # 2.4.179 Note
  # The Note record specifies a comment associated with a cell or revision information about a comment associated with a cell.
  # @return [Hash]
  def r_note
    @serial = true

    Unxls::Biff8::Structure.notesh(@bytes)
  end

  # 2.4.180 Number
  # The Number record specifies a cell that contains a floating-point number.
  # @return [Hash]
  def r_number
    @serial = true

    {
      **Unxls::Biff8::Structure.cell(@bytes.read(6)), # cell (6 bytes): A Cell structure that specifies the cell.
      num: Unxls::Biff8::Structure.xnum(@bytes.read(8)) # num (8 bytes): An Xnum (section 2.5.342) value that specifies the cell value. @todo ChartNumNillable
    }
  end

  # 2.4.181 Obj
  # The Obj record specifies the properties of an object in a sheet.
  # @todo May be Continued-ed (!!!)
  # @return [Hash]
  def r_obj
    @serial = true

    {
      cmo: Unxls::Biff8::Structure.ftcmo(@bytes.read(22)), # cmo (22 bytes): An FtCmo structure that specifies the common properties of this object.
      _other_fields: :not_implemented # @todo
    }
  end

  # @todo May be Continue-ed (!!!)
  # 2.4.192 PhoneticInfo
  # The PhoneticInfo record specifies the default format for phonetic strings and the ranges of cells on the sheet that have phonetic strings that are visible.
  # @return [Hash]
  def r_phoneticinfo
    {
      phs: Unxls::Biff8::Structure.phs(@bytes.read(4)), # phs (4 bytes): A Phs structure that specifies the default format for phonetic strings on the sheet. When a phonetic string is entered into a cell that does not already contain a phonetic string, the default format is applied to the phonetic string.
      sqref: Unxls::Biff8::Structure.sqref(@bytes) # sqref (variable): An SqRef structure that specifies the ranges of cells on the sheet that have phonetic strings that are visible.
    }
  end

  # 2.4.220 RK
  # The RK record specifies the numeric data contained in a single cell.
  # @return [Hash]
  def r_rk
    @serial = true

    rw, col = @bytes.read(4).unpack('vv')

    {
      rw: rw, # rw (2 bytes): An Rw structure that specifies a row index.
      col: col, # col (2 bytes): A Col structure that specifies a column index.
      **Unxls::Biff8::Structure.rkrec(@bytes.read(6)) # rkrec (6 bytes): An RkRec structure that specifies the numeric data for a single cell.
    }
  end

  # 2.4.221 Row, page 377
  # The Row record specifies a single row on a sheet.
  # @return [Hash]
  def r_row
    @serial = true

    rw, col_mic, col_mac, miy_rw, _, attrs = @bytes.read.unpack('vvvvVV')
    attrs = Unxls::BitOps.new(attrs)

    {
      rw: rw, # rw (2 bytes): A Rw (2.5.227) structure that specifies the row index.
      colMic: col_mic, # colMic (2 bytes): An unsigned integer that specifies the zero-based index of the first column that contains a cell populated with data or formatting in the current row. MUST be less than or equal to 255.
      colMac: col_mac, # colMac (2 bytes): An unsigned integer that specifies the one-based index of the last column that contains a cell populated with data or formatting in the current row. MUST be less than or equal to 256. If colMac is equal to colMic, this record specifies a row with no CELL records.
      miyRw: miy_rw, # miyRw (2 bytes): An unsigned integer that specifies the row height in twips. If fDyZero is 1, the row is hidden and the value of miyRw specifies the original row height. MUST be greater than or equal to 2 and MUST be less than or equal to 8192.
      # reserved1 (2 bytes): MUST be zero, and MUST be ignored.
      # unused1 (2 bytes): Undefined and MUST be ignored.
      iOutLevel: attrs.value_at(0..2), # A - iOutLevel (3 bits): An unsigned integer that specifies the outline level of the row.
      # 3 B - reserved2 (1 bit): MUST be zero, and MUST be ignored.
      fCollapsed: attrs.set_at?(4), # C - fCollapsed (1 bit): A bit that specifies whether the rows that are one level of outlining deeper than the current row are included in the collapsed outline state.
      fDyZero: attrs.set_at?(5), # D - fDyZero (1 bit): A bit that specifies whether the row is hidden.
      fUnsynced: attrs.set_at?(6), # E - fUnsynced (1 bit): A bit that specifies whether the row height was manually set.
      fGhostDirty: attrs.set_at?(7), # F - fGhostDirty (1 bit): A bit that specifies whether the row was formatted.
      # 8…15 reserved3 (1 byte): MUST be 1, and MUST be ignored.
      ixfe: attrs.value_at(16..27), # index of XF record of row formatting
      fExAsc: attrs.set_at?(28), # G - fExAsc (1 bit): A bit that specifies whether any cell in the row has a thick top border, or any cell in the row directly above the current row has a thick bottom border. Thick borders are specified by the following enumeration values from BorderStyle: THICK and DOUBLE.
      fExDes: attrs.set_at?(29), # H - fExDes (1 bit): A bit that specifies whether any cell in the row has a medium or thick bottom border, or any cell in the row directly below the current row has a medium or thick top border. Thick borders are previously specified. Medium borders are specified by the following enumeration values from BorderStyle: MEDIUM, MEDIUMDASHED, MEDIUMDASHDOT, MEDIUMDASHDOTDOT, and SLANTDASHDOT.
      fPhonetic: attrs.set_at?(30), # I - fPhonetic (1 bit): A bit that specifies whether the phonetic guide feature is enabled for any cell in this row.
      # J - unused2 (1 bit): Undefined and MUST be ignored.
    }
  end

  # @todo May be Continue-ed (!!!)
  # 2.4.268 String
  # The String record specifies the string value of a formula (section 2.2.2).
  # @return [Hash]
  def r_string
    @serial = true

    {
      string: Unxls::Biff8::Structure.xlunicodestring(@bytes.read) # string (variable): An XLUnicodeString structure that specifies the string value of a formula (section 2.2.2). The value of string.cch MUST be less than or equal to 32767.
    }
  end

  # 2.4.313 SxView
  # The SxView record specifies PivotTable view information and that specifies the beginning of a collection of records as defined by the Worksheet substream ABNF. The collection specifies the remainder of the PivotTable view.
  # @return [Hash]
  def r_sxview
    @serial = true

    ref = Unxls::Biff8::Structure.ref8u(@bytes.read(8))
    rw_first_head, rw_first_data, col_first_data, i_cache, _, sxaxis4_data = @bytes.read(12).unpack('vvvs<vv')

    result = {
      ref: ref, # ref (8 bytes): A Ref8U structure that specifies the PivotTable report body. For more information, see Location and Body.
      rwFirstHead: rw_first_head, # rwFirstHead (2 bytes): An RwU structure that specifies the first row of the row area. MUST be 1 if none of the axes are assigned in this PivotTable view. Otherwise, the value MUST be greater than or equal to ref.rwFirst.
      rwFirstData: rw_first_data, # rwFirstData (2 bytes): An RwU structure that specifies the first row of the data area. MUST be 1 if none of the axes are assigned in this PivotTable view. Otherwise, it MUST be equal to the value as specified by the following formula: rwFirstData = rwFirstHead + cDimCol
      colFirstData: col_first_data, # colFirstData (2 bytes): A ColU structure that specifies the first column of the data area. It MUST be 1 if none of the axes are assigned in this PivotTable view. Otherwise, the value MUST be greater than or equal to ref.colFirst, and if the value of cDimCol or cDimData is not zero, it MUST be less than or equal to ref.colLast.
      iCache: i_cache, # iCache (2 bytes): A signed integer that specifies the zero-based index of an SXStreamID record in the Globals Substream. See Associated PivotCache for more information. MUST be greater than or equal to zero and less than the number of SXStreamID records in the Globals Substream.
      # reserved (2 bytes): MUST be zero, and MUST be ignored.
      sxaxis4Data: Unxls::Biff8::Structure.sxaxis4data(sxaxis4_data) # sxaxis4Data (2 bytes): An SXAxis structure that specifies the default axis for the data field. Either the sxaxis4Data.sxaxisRw field MUST be 1 or the sxaxis4Data.sxaxisCol field MUST be 1. The sxaxis4Data.sxaxisPage field MUST be 0 and the sxaxis4Data.sxaxisData field MUST be 0.
    }
    ipos4_data, c_dim, c_dim_rw, c_dim_col, c_dim_pg, c_dim_data, c_rw, c_col = @bytes.read(16).unpack('s<2v3s<v2')

    result.merge!({
      ipos4Data: ipos4_data, # ipos4Data (2 bytes): A signed integer that specifies the row or column position for the data field in the PivotTable view. The sxaxis4Data field specifies whether this is a row or column position. MUST be greater than or equal to -1 and less than or equal to 0x7FFF. A value of -1 specifies the default position.
      cDim: c_dim, # cDim (2 bytes): A signed integer that specifies the number of pivot fields in the PivotTable view. MUST equal the number of Sxvd records following this record. MUST equal the number of fields in the associated PivotCache specified by iCache.
      cDimRw: c_dim_rw, # cDimRw (2 bytes): An unsigned integer that specifies the number of fields on the row axis of the PivotTable view. MUST be less than or equal to 0x7FFF. MUST equal the number of array elements in the SxIvd record in this PivotTable view that contain row items.
      cDimCol: c_dim_col, # cDimCol (2 bytes): An unsigned integer that specifies the number of fields on the column axis of the PivotTable view. MUST be less than or equal to 0x7FFF. MUST equal the number of array elements in the SxIvd record in this PivotTable view that contain column items.
      cDimPg: c_dim_pg, # cDimPg (2 bytes): An unsigned integer that specifies the number of page fields in the PivotTable view. MUST be less than or equal to 0x7FFF. MUST equal the number of array elements in the SXPI record in this PivotTable view.
      cDimData: c_dim_data, # cDimData (2 bytes): A signed integer that specifies the number of data fields in the PivotTable view. MUST be greater than or equal to zero and less than or equal to 0x7FFF. MUST equal the number of SXDI records in this PivotTable view.
      cRw: c_rw, # cRw (2 bytes): An unsigned integer that specifies the number of pivot lines in the row area of the PivotTable view. MUST be less than or equal to 0x7FFF. MUST equal the number of array elements in the first SXLI record in this PivotTable view.
      cCol: c_col, # cCol (2 bytes): An unsigned integer that specifies the number of pivot lines in the column area of the PivotTable view. MUST equal the number of array elements in the second SXLI record in this PivotTable view.
    })

    attrs = Unxls::BitOps.new(@bytes.read(2).unpack('v').first)

    result.merge!({
      fRwGrand: attrs.set_at?(0), # A - fRwGrand (1 bit): A bit that specifies whether the PivotTable contains grand totals for rows. MUST be 0 if none of the axes have been assigned in this PivotTable view.
      fColGrand: attrs.set_at?(1), # B - fColGrand (1 bit): A bit that specifies whether the PivotTable contains grand totals for columns. MUST be 1 if none of the axes are assigned in this PivotTable view.
      # C - unused1 (1 bit):  Undefined and MUST be ignored.
      fAutoFormat: attrs.set_at?(3), # D - fAutoFormat (1 bit): A bit that specifies whether the PivotTable has AutoFormat applied.
      fAtrNum: attrs.set_at?(4), # E - fAtrNum (1 bit): A bit that specifies whether the PivotTable has number AutoFormat applied.
      fAtrFnt: attrs.set_at?(5), # F - fAtrFnt (1 bit): A bit that specifies whether the PivotTable has font AutoFormat applied.
      fAtrAlc: attrs.set_at?(6), # G - fAtrAlc (1 bit): A bit that specifies whether the PivotTable has alignment AutoFormat applied.
      fAtrBdr: attrs.set_at?(7), # H - fAtrBdr (1 bit): A bit that specifies whether the PivotTable has border AutoFormat applied.
      fAtrPat: attrs.set_at?(8), # I - fAtrPat (1 bit): A bit that specifies whether the PivotTable has pattern AutoFormat applied.
      fAtrProc: attrs.set_at?(9), # J - fAtrProc (1 bit): A bit that specifies whether the PivotTable has width/height AutoFormat applied.
      # unused2 (6 bits): Undefined and MUST be ignored.
    })

    itbl_auto_fmt, cch_table_name, cch_data_name = @bytes.read(6).unpack('vvv')

    result.merge!({
      itblAutoFmt: Unxls::Biff8::Structure.autofmt8(itbl_auto_fmt), # itblAutoFmt (2 bytes): An AutoFmt8 structure that specifies the PivotTable AutoFormat. If the value of itblAutoFmt in the associated SXViewEx9 record is not 1, this field is overridden by the value of itblAutoFmt in the associated SXViewEx9.
      cchTableName: cch_table_name, # cchTableName (2 bytes): An unsigned integer that specifies the length, in characters, of stTable. MUST be greater than or equal to zero and less than or equal to 0x00FF.
      cchDataName: cch_data_name, # cchDataName (2 bytes): An unsigned integer that specifies the length, in characters of stData. MUST be greater than zero and less than or equal to 0x00FE.
      stTable: Unxls::Biff8::Structure.xlunicodestringnocch(@bytes, cch_table_name), # stTable (variable): An XLUnicodeStringNoCch structure that specifies the name of the PivotTable. The length of this field is specified by cchTableName.
      stData: Unxls::Biff8::Structure.xlunicodestringnocch(@bytes, cch_data_name) # stData (variable): An XLUnicodeStringNoCch structure that specifies the name of the data field. The length of this field is specified by cchDataName.
    })
  end

  # 2.4.315 SXViewEx9
  # The SXViewEx9 record specifies extensions to the PivotTable view.
  # @return [Hash]
  def r_sxviewex9
    @serial = true

    result = { frtHeader: Unxls::Biff8::Structure.frtheader(@bytes.read(8)) }

    attrs, itbl_auto_fmt = @bytes.read(6).unpack('Vv')
    attrs = Unxls::BitOps.new(attrs)
    result.merge!({
      # C - reserved4 (1 bit): MUST be zero, and MUST be ignored.
      fPrintTitles: attrs.set_at?(1), # D - fPrintTitles (1 bit): A bit that specifies whether the print titles for the worksheet are set based on the PivotTable report. The row print titles are set to the pivot item captions on the column axis and the column print titles are set to the pivot item captions on the row axis.
      fLineMode: attrs.set_at?(2), # E - fLineMode (1 bit): A bit that specifies whether any pivot field is in outline mode. See Subtotalling for more information.
      # F - reserved5 (2 bits): MUST be zero, and MUST be ignored.
      fRepeatItemsOnEachPrintedPage: attrs.set_at?(5), # G - fRepeatItemsOnEachPrintedPage (1 bit): A bit that specifies whether pivot item captions on the row axis are repeated on each printed page for pivot fields in tabular form.
      # reserved6 (26 bits): MUST be zero, and MUST be ignored.
      itblAutoFmt: Unxls::Biff8::Structure.autofmt8(itbl_auto_fmt), # itblAutoFmt (2 bytes): An AutoFmt8 structure that specifies the PivotTable AutoFormat. If the value of this field is not 1, this field overrides the itblAutoFmt field in the previous SxView record.
      chGrand: Unxls::Biff8::Structure.xlunicodestring(@bytes) # chGrand (variable): An XLUnicodeString  structure that specifies a user-entered caption to display for grand totals when the PivotTable is recalculated. The length MUST be less than or equal to 255 characters.
    })
  end

  # 2.4.329 TxO
  # @todo Check: may be Continued-ed
  # The TxO record specifies the text in a text box or a form control.
  # @return [Hash]
  def r_txo
    @serial = true

    attrs, rot = @bytes.read(4).unpack('vv')
    attrs = Unxls::BitOps.new(attrs)

    # A - reserved1 (1 bit): MUST be zero, and MUST be ignored.

    h_alignment = attrs.value_at(1..3) # B - hAlignment (3 bits): An unsigned integer that specifies the horizontal alignment.
    h_alignment_d = {
      1 => :left,
      2 => :centered,
      3 => :right,
      4 => :justify,
      7 => :justify_distributed
    }.freeze[h_alignment]
    
    v_alignment = attrs.value_at(4..6) # C - vAlignment (3 bits): An unsigned integer that specifies the vertical alignment.
    v_alignment_d = {
      1 => :top,
      2 => :middle,
      3 => :bottom,
      4 => :justify,
      7 => :justify_distributed
    }.freeze[v_alignment]
      
    result = {
      hAlignment: h_alignment,
      hAlignment_d: h_alignment_d,
      vAlignment: v_alignment,
      vAlignment_d: v_alignment_d,
      # D - reserved2 (2 bits): MUST be zero, and MUST be ignored.
      fLockText: attrs.set_at?(9), # E - fLockText (1 bit): A bit that specifies whether the text is locked.
      # F - reserved3 (4 bits): MUST be zero, and MUST be ignored.
      fJustLast: attrs.set_at?(14), # G - fJustLast (1 bit): A bit that specifies whether the justify alignment or justify distributed alignment is used on the last line of the text in specific versions of the application.
      fSecretEdit: attrs.set_at?(15), # H - fSecretEdit (1 bit): A bit that specifies whether this is a text box used for typing passwords and hiding the actual characters being typed by the user.
    }

    result[:rot] = rot # rot (2 bytes): An unsigned integer that specifies the orientation of the text within the object boundary.
    result[:rot_d] = {
      0 => :none, # Specifies no rotation
      1 => :stacked, # Specifies stacked or vertical orientation
      2 => :ccw_90, # Specifies 90-degree counter-clockwise rotation
      3 => :cw_90, # Specifies 90-degree clockwise rotation
    }.freeze[rot]

    control_obj_types = Set[0, 5, 7, 11, 12, 14].freeze # Group, Chart, Button, Checkbox, Radio button, Label

    if (preceding_obj_record = @stream.last_parsed[:Obj])
      @bytes.read(6) # reserved4 (2 bytes), reserved5 (4 bytes): MUST be zero and MUST be ignored. This field MUST exist if and only if the value of cmo.ot in the preceding Obj record is not 0, 5, 7, 11, 12 or 14.
      if control_obj_types.include?(preceding_obj_record.last[:cmo][:ot])
        result[:controlInfo] = :not_implemented # controlInfo (6 bytes): An optional ControlInfo (2.5.61) structure that specifies the properties for some form controls. The field MUST exist if and only if the value of cmo.ot in the preceding Obj record is 0, 5, 7, 11, 12, or 14.
      end
    end

    result[:cchText] = @bytes.read(2).unpack('v').first # cchText (2 bytes): An unsigned integer that specifies the number of characters in the text string contained in the Continue records immediately following this record.
    result[:cbRuns] = @bytes.read(2).unpack('v').first # cbRuns (2 bytes): An unsigned integer that specifies the number of bytes of formatting run information in the TxORuns structure contained in the Continue records following this record. If cchText is 0, this value MUST be 0. Otherwise, the value MUST be greater than or equal to 16 and MUST be a multiple of 8.
    result[:ifntEmpty] = @bytes.read(2).unpack('v').first # ifntEmpty (2 bytes): A FontIndex structure that specifies the font when the value of cchText is 0.

    @bytes.read # skip the rest of the record
    result[:fmla] = :not_implemented # fmla (variable): An ObjFmla (2.5.187) structure that specifies the parsed expression of the formula (section 2.2.2) for the text.

    # see 2.5.296 XLUnicodeStringNoCch
    if result[:cchText] > 0
      open_next_record_block
      high_byte = Unxls::BitOps.new(@bytes.read(1).unpack('C').first).set_at?(0)
      result[:text_string] = Unxls::Biff8::Structure._read_continued_string(self, result[:cchText], high_byte)
    end

    # see 2.5.272 TxORuns
    num_of_formatting_runs = result[:cbRuns] / 8 - 1
    if num_of_formatting_runs > 0
      result[:formatting_runs] = []
      num_of_formatting_runs.times do
        open_next_record_block if @bytes.eof?
        result[:formatting_runs] << Unxls::Biff8::Structure.run(@bytes)
        break if end_of_data?
      end
    end

    result
  end

  # 2.4.351 WsBool
  # The WsBool record specifies information about a sheet.
  # @return [Hash]
  def r_wsbool
    attrs = @bytes.read.unpack('v').first
    attrs = Unxls::BitOps.new(attrs)

    {
      fShowAutoBreaks: attrs.set_at?(0), # A - fShowAutoBreaks (1 bit): A bit that specifies whether page breaks inserted automatically are visible on the sheet.
      # B - reserved1 (3 bits): MUST be zero, and MUST be ignored.
      fDialog: attrs.set_at?(4), # C - fDialog (1 bit): A bit that specifies whether the sheet is a dialog sheet.
      fApplyStyles: attrs.set_at?(5), # D - fApplyStyles (1 bit): A bit that specifies whether to apply styles in an outline when an outline is applied.
      fRowSumsBelow: attrs.set_at?(6), # E - fRowSumsBelow (1 bit): A bit that specifies whether summary rows appear below an outline's detail rows.
      fColSumsRight: attrs.set_at?(7), # F - fColSumsRight (1 bit): A bit that specifies whether summary columns appear to the right or left of an outline's detail columns.
      fFitToPage: attrs.set_at?(8), # G - fFitToPage (1 bit): A bit that specifies whether to fit the printable contents to a single page when printing this sheet.
      # H - reserved2 (1 bit): MUST be zero, and MUST be ignored.
      # I - unused (2 bits): Undefined and MUST be ignored.
      fSyncHoriz: attrs.set_at?(12), # J - fSyncHoriz (1 bit): A bit that specifies whether horizontal scrolling is synchronized across multiple windows displaying this sheet.
      fSyncVert: attrs.set_at?(13), # K - fSyncVert (1 bit): A bit that specifies whether vertical scrolling is synchronized across multiple windows displaying this sheet.
      fAltExprEval: attrs.set_at?(14), # L - fAltExprEval (1 bit): A bit that specifies whether the sheet uses transition formula evaluation.
      fFormulaEntry: attrs.set_at?(15), # M - fAltFormulaEntry (1 bit): A bit that specifies whether the sheet uses transition formula entry.
    }
  end

end
