# frozen_string_literal: true

# 2.5.198 Parsed Expressions
module Unxls::Biff8::ParsedExpressions
  extend self

  # 2.5.198.2 BErr
  # @param id [Integer]
  # @return [Symbol]
  def berr(id)
    {
      0x00 => :'NULL!',
      0x07 => :'DIV/0!',
      0x0F => :'VALUE!',
      0x17 => :'REF!',
      0x1D => :'NAME?',
      0x24 => :'NUM!',
      0x2A => :'N/A',
    }[id]
  end
end

# Structure-processing methods
#
# +method naming+
# processing methods for structures from the 2.5 list: <structure name downcase>, i.e. 'cellxf'
# helper methods for data structures not from the specification list: _<method name> i.e. '_encode_string'
module Unxls::Biff8::Structure
  using Unxls::Helpers

  extend self
  extend Unxls::Biff8::ParsedExpressions

  # 2.5.9 AutoFmt8
  # The AutoFmt8 enumeration specifies the following auto formatting styles.
  # @param value [Integer]
  # @return [Symbol]
  def autofmt8(value)
    {
      0x0000 => :XL8_ITBLSIMPLE, # Simple
      0x0001 => :XL8_ITBLCLASSIC1, # Classic 1
      0x0002 => :XL8_ITBLCLASSIC2, # Classic 2
      0x0003 => :XL8_ITBLCLASSIC3, # Classic 3
      0x0004 => :XL8_ITBLACCOUNTING1, # Accounting 1
      0x0005 => :XL8_ITBLACCOUNTING2, # Accounting 2
      0x0006 => :XL8_ITBLACCOUNTING3, # Accounting 3
      0x0007 => :XL8_ITBLACCOUNTING4, # Accounting 4
      0x0008 => :XL8_ITBLCOLORFUL1, # Colorful 1
      0x0009 => :XL8_ITBLCOLORFUL2, # Colorful 2
      0x000A => :XL8_ITBLCOLORFUL3, # Colorful 3
      0x000B => :XL8_ITBLLIST1, # List 1
      0x000C => :XL8_ITBLLIST2, # List 2
      0x000D => :XL8_ITBLLIST3, # List 3
      0x000E => :XL8_ITBL3DEFFECTS1, # 3Deffects 1
      0x000F => :XL8_ITBL3DEFFECTS2, # 3Deffects 2
      0x0010 => :XL8_ITBLNONE_GEN, # None
      0x0011 => :XL8_ITBLJAPAN2, # Japan 2
      0x0012 => :XL8_ITBLJAPAN3, # Japan 3
      0x0013 => :XL8_ITBLJAPAN4, # Japan 4
      0x0014 => :XL8_ITBLNONE_JPN, # Japan None
      0x1000 => :XL8_ITBLREPORT1, # Report 1
      0x1001 => :XL8_ITBLREPORT2, # Report 2
      0x1002 => :XL8_ITBLREPORT3, # Report 3
      0x1003 => :XL8_ITBLREPORT4, # Report 4
      0x1004 => :XL8_ITBLREPORT5, # Report 5
      0x1005 => :XL8_ITBLREPORT6, # Report 6
      0x1006 => :XL8_ITBLREPORT7, # Report 7
      0x1007 => :XL8_ITBLREPORT8, # Report 8
      0x1008 => :XL8_ITBLREPORT9, # Report 9
      0x1009 => :XL8_ITBLREPORT10, # Report 10
      0x100A => :XL8_ITBLTABLE1, # Table 1
      0x100B => :XL8_ITBLTABLE2, # Table 2
      0x100C => :XL8_ITBLTABLE3, # Table 3
      0x100D => :XL8_ITBLTABLE4, # Table 4
      0x100E => :XL8_ITBLTABLE5, # Table 5
      0x100F => :XL8_ITBLTABLE6, # Table 6
      0x1010 => :XL8_ITBLTABLE7, # Table 7
      0x1011 => :XL8_ITBLTABLE8, # Table 8
      0x1012 => :XL8_ITBLTABLE9, # Table 9
      0x1013 => :XL8_ITBLTABLE10, # Table 10
      0x1014 => :XL8_ITBLPTCLASSIC, # Table PTClassic
      0x1015 => :XL8_ITBLPTNONE, # None
    }[value]
  end

  # 2.5.10 Bes
  # The Bes structure specifies either a Boolean (section 2.5.14) value or an error value.
  # @param data [String]
  # @return [Hash]
  def bes(data)
    b_bool_err, f_error = data.unpack('CC')

    b_bool_err_d = case f_error
    when 0
      {
        0x00 => :False,
        0x01 => :True
      }

    when 1
      {
        0x00 => :'#NULL!',
        0x07 => :'#DIV/0!',
        0x0F => :'#VALUE!',
        0x17 => :'#REF!',
        0x1D => :'#NAME?',
        0x24 => :'#NUM!',
        0x2A => :'#N/A',
        0x2B => :'#GETTING_DATA',
      }

    else
      raise "Unexpected fError value #{f_error} in Bes structure"

    end

    {
      bBoolErr: b_bool_err, # bBoolErr (1 byte): An unsigned integer that specifies either a Boolean value or an error value, depending on the value of fError.
      bBoolErr_d: b_bool_err_d[b_bool_err],
      fError: f_error # fError (1 byte):  A Boolean that specifies whether bBoolErr contains an error code or a Boolean value.
    }
  end

  # @param data [String]
  # @return [true, false]
  def _bool1b(data)
    _int1b(data) == 1
  end

  # 2.5.11 Bold,
  # Also see 2.5.248 Stxp
  # @param id [Integer]
  # @return [Symbol]
  def bold(id)
    {
      0x0190 => :BLSNORMAL, # Normal font weight
      0x02BC => :BLSBOLD, # Bold font weight
      0xFFFF => :ignored # Indicates that this specification is to be ignored
    }[id]
  end

  # 2.5.15 BorderStyle
  # @param id [Integer]
  # @return [Symbol]
  def borderstyle(id)
    {
      0x0000 => :NONE, # No border
      0x0001 => :THIN, # Thin line
      0x0002 => :MEDIUM, # Medium line
      0x0003 => :DASHED, # Dashed line
      0x0004 => :DOTTED, # Dotted line
      0x0005 => :THICK, # Thick line
      0x0006 => :DOUBLE, # Double line
      0x0007 => :HAIR, # Hairline
      0x0008 => :MEDIUMDASHED, # Medium dashed line
      0x0009 => :DASHDOT, # Dash-dot line
      0x000A => :MEDIUMDASHDOT, # Medium dash-dot line
      0x000B => :DASHDOTDOT, # Dash-dot-dot line
      0x000C => :MEDIUMDASHDOTDOT, # Medium dash-dot-dot line
      0x000D => :SLANTEDDASHDOTDOT, # Slanted dash-dot-dot line
    }[id]
  end

  # 2.5.16 BuiltInStyle
  # The BuiltInStyle structure specifies the type of a built-in cell style. For row outline and column outline types this structure also specifies the outline level of the style.
  # @param data [String]
  # @return [Hash]
  def builtinstyle(data)
    isty_built_in, i_level = data.unpack('CC')

    name_base = Unxls::Biff8::Constants::BUILTIN_STYLES[isty_built_in]
    name = (1..2).include?(isty_built_in) ? "#{name_base}#{i_level + 1}".to_sym : name_base

    {
      istyBuiltIn: isty_built_in, # istyBuiltIn (1 byte): An unsigned integer that specifies the type of the built-in cell style.
      iLevel: i_level, # iLevel (1 byte): An unsigned integer that specifies the depth level of row/column automatic outlining.
      istyBuiltIn_d: name
    }
  end

  # See 2.5.282 XFIndex, p. 885
  # @param value [Integer]
  # @return [Symbol]
  def _builtin_xf_description(value)
    {
      0 => :'Normal style', # fStyle value == 1
      1 => :'Row outline level 1', # 1
      2 => :'Row outline level 2', # 1
      3 => :'Row outline level 3', # 1
      4 => :'Row outline level 4', # 1
      5 => :'Row outline level 5', # 1
      6 => :'Row outline level 6', # 1
      7 => :'Row outline level 7', # 1
      8 => :'Column outline level 1', # 1
      9 => :'Column outline level 2', # 1
      10 => :'Column outline level 3', # 1
      11 => :'Column outline level 4', # 1
      12 => :'Column outline level 5', # 1
      13 => :'Column outline level 6', # 1
      14 => :'Column outline level 7', # 1
      15 => :'Default cell format', # 0
    }[value]
  end

  # 2.5.17 CachedDiskHeader
  # The CachedDiskHeader structure specifies the formatting information of a table column heading.
  # @param io [StringIO]
  # @param f_save_style_name [true, false]
  # @return [Hash]
  def cacheddiskheader(io, f_save_style_name)
    cbdxf_hdr_disk = io.read(4).unpack('V').first

    result = {
      cbdxfHdrDisk: cbdxf_hdr_disk, # cbdxfHdrDisk (4 bytes): An unsigned integer that specifies the size, in bytes, of the rgHdrDisk field.
      rgHdrDisk: dxfn12list(io.read(cbdxf_hdr_disk)) # rgHdrDisk (variable): A DXFN12List structure that specifies the formatting of the column heading.
    }

    # If present, the formatting as specified by strStyleName is applied first, before the formatting as specified by rgHdrDisk is applied.
    result[:strStyleName] = xlunicodestring(io) if f_save_style_name # strStyleName (variable): An XLUnicodeString that specifies the name of the style to use for the column heading. The name of the style MUST equal the user field of a Style record in the Globals Substream ABNF, or the name of a built-in style, as specified by the BuiltInStyle record. This field is present only if the fSaveStyleName field of the containing Feat11FieldDataItem structure is set to 0x1.

    result
  end

  # 2.5.19 Cell
  # The Cell structure specifies a cell in the current sheet.
  # @param data [String]
  # @return [Hash]
  def cell(data)
    rw, col, ixfe = data.unpack('v3')

    {
      rw: rw, # rw (2 bytes): An Rw that specifies the row.
      col: col, # col (2 bytes): A Col that specifies the column.
      ixfe: ixfe # ixfe (2 bytes): An IXFCell that specifies the XF record.
    }
  end

  # 2.5.20 CellXF, p. 597
  # See also 2.5.91 DXFALC
  # This structure specifies formatting properties for a cell.
  # @param io [StringIO]
  # @return [Hash]
  def cellxf(io)
    result = _xfalc(io)
    result.delete(:fMergeCell)

    result.merge!(_xfbdr(io))

    attrs = io.read(2).unpack('v').first
    attrs = Unxls::BitOps.new(attrs)
    result[:icvFore] = attrs.value_at(0..6) # icvFore (7 bits): An IcvXF that specifies the foreground color of the fill pattern.
    result[:icvBack] = attrs.value_at(7..13) # icvBack (7 bits): An unsigned integer that specifies the background color of the fill pattern.
    result[:fsxButton] = attrs.set_at?(14) # fsxButton (1 bit): A bit that specifies whether the XF record is attached to a pivot field drop-down button.
    # Q - reserved3 (1 bit): MUST be 0 and MUST be ignored.

    result
  end

  # 2.5.27 CFExNonCF12
  # The CFExNonCF12 structure specifies properties that extend a conditional formatting rule that is specified by a CF record.
  # @param io [StringIO]
  # @return [Hash]
  def cfexnoncf12(io)
    icf, cp, icf_template, ipriority, a_e_raw, f_has_dxf = io.read(8).unpack('vCCvCC')
    a_e = Unxls::BitOps.new(a_e_raw)
    result = {
      icf: icf, # icf (2 bytes): An unsigned integer that specifies a zero-based index of a CF record in the collection of CF records directly following the CondFmt record that is referenced by the parent CFEx record with the nID field. The referenced CF specifies the conditional formatting rule to be extended.
      cp: cp, # cp (1 byte): An unsigned integer that specifies the type of comparison operation to use when the ct field of the CF record referenced by the icf field of this structure is equal to 0x01.
      cp_d: _cf_cp(cp),
      icfTemplate: icf_template, # icfTemplate (1 byte): An unsigned integer that specifies the template from which the rule was created. MUST be the least significant byte of one of the valid values specified for the icfTemplate field in the CF12 record.
      icfTemplate_d: _icf_template_d(icf_template),
      ipriority: ipriority, # ipriority (2 bytes): An unsigned integer that specifies the priority of the rule. Rules that apply to the same cell are evaluated in increasing order of ipriority. MUST be unique across all CF12 records and CFExNonCF12 structures in the worksheet substream.
      fActive: a_e.set_at?(0), # A - fActive (1 bit): A bit that specifies whether the rule is active. If set to zero, the rule will be ignored.
      fStopIfTrue: a_e.set_at?(1), # B - fStopIfTrue (1 bit): A bit that specifies whether, when a cell fulfills the condition corresponding to this rule, the lower priority conditional formatting rules that apply to this cell are evaluated.
      # C - reserved1 (1 bit): MUST be zero and MUST be ignored.
      # D - unused (1 bit): Undefined and MUST be ignored.
      # E - reserved2 (4 bits): MUST be zero and MUST be ignored.
      fHasDXF: f_has_dxf, # fHasDXF (1 byte): A Boolean (section 2.5.14) that specifies whether cell formatting data is part of this record extension.
    }

    result[:dxf] = dxfn12(io) if f_has_dxf == 1

    cb_template_parm = io.read(1).unpack('C').first
    result[:cbTemplateParm] = cb_template_parm # cbTemplateParm (1 byte): An unsigned integer that specifies the size of the rgbTemplateParms field in bytes. MUST be 16.
    result[:rgbTemplateParms] = cfextemplateparams(io.read(cb_template_parm), icf_template) # rgbTemplateParms (16 bytes): A CFExTemplateParams that specifies the parameters for the rule.

    result
  end

  # 2.5.21 CFColor
  # The CFColor structure specifies a color in conditional formatting records or in a SheetExt record.
  # @param io [StringIO]
  # @return [Hash]
  def cfcolor(io)
    xclr_type = io.read(4).unpack('V').first
    type = xcolortype(xclr_type)

    {
      xclrType: xclr_type, # xclrType (4 bytes): An XColorType that specifies the type of color reference.
      xclrType_d: type,
      xclrValue: _interpret_xclr(type, io.read(4)), # xclrValue (4 bytes): A structure that specifies the color value.
      numTint: xnum(io.read(8)) # numTint (8 bytes): An Xnum (section 2.5.342) that specifies the tint and shade value to be applied to the color.
    }
  end

  # See cp value meaning in 2.4.43 CF12 and 2.4.42 CF
  # @param value [Integer]
  # @return [Hash]
  def _cf_cp(value)
    {
      0x00 => :'No comparison',
      # v between v1 and v2 or strictly
      # v2 is greater than or equal to v1, and v is greater than or equal to v1 and less than or equal to v2 –Or–
      # v1 is greater than v2, and v is greater than or equal to v2 and less than or equal to v1
      0x01 => :'((v1 <= v2) && (v1 <= v <= v2)) || ((v2 < v1) && (v1 >= v >= v2))',
      # v is not between v1 and v2 or strictly
      # v2 is greater than or equal to v1, and v is less than v1 or greater than v2 –Or–
      # v1 is greater than v2, and v is less than v2 or greater than v1
      0x02 => :'((v1 <= v2) && (v < v1 || v > v2)) || ((v1 > v2) && (v > v1 || v < v2))',
      0x03 => :'v == v1', # v is equal to v1
      0x04 => :'v != v1', # v is not equal to v1
      0x05 => :'v > v1', # v is greater than v1
      0x06 => :'v < v1', # v is less than v1
      0x07 => :'v >= v1', # v is greater than or equal to v1
      0x08 => :'v <= v1' # v is less than or equal to v1
    }[value]
  end

  # 2.5.22 CFDatabar
  # The CFDatabar structure specifies the parameters of a conditional formatting rule that uses data bar formatting.
  # @param io [StringIO]
  # @return [Hash]
  def cfdatabar(io)
    _, _, a_r2_raw, i_percent_min, i_percent_max = io.read(6).unpack('vCCCC')
    a_r2 = Unxls::BitOps.new(a_r2_raw)

    {
      _type: :CFDatabar,
      # unused (2 bytes): Undefined and MUST be ignored.
      # reserved1 (1 byte): MUST be zero and MUST be ignored.
      fRightToLeft: a_r2.set_at?(0), # A - fRightToLeft (1 bit): A bit that specifies whether the data bars are drawn starting from the right of the cell.
      fShowValue: a_r2.set_at?(1), # B - fShowValue (1 bit): A bit that specifies whether the numerical value of the cell appears in the cell along with the data bar.
      # reserved2 (6 bits): MUST be zero and MUST be ignored.
      iPercentMin: i_percent_min, # iPercentMin (1 byte): An unsigned integer that specifies the length of a data bar, as a percentage of the cell width, that is applied to cells with values equal to the CFVO value specified by cfvoDB1.
      iPercentMax: i_percent_max, # iPercentMax (1 byte): An unsigned integer that specifies the length of a data bar, as a percentage of the cell width, that is applied to cells with values equal to the CFVO value specified by cfvoDB2.
      color: cfcolor(io), # color (16 bytes): A CFColor structure that specifies the color of the data bar.
      cfvoDB1: cfvo(io), # cfvoDB1 (variable): A CFVO that specifies the maximum cell value that will be represented with a minimum width data bar. All cell values that are less than or equal to the CFVO value specified by this field are represented with a data bar of iPercentMin percent of the cell width.
      cfvoDB2: cfvo(io) # cfvoDB2 (variable): A CFVO that specifies the minimum cell value that will be represented with a maximum width data bar. All cell values that are greater than or equal to the CFVO value specified by this field are represented with a data bar of iPercentMax percent of the cell width.
    }
  end

  # 2.5.23 CFExAveragesTemplateParams
  # This structure specifies the parameters for an above or below average conditional formatting rule in a containing CF12 record or CFExNonCF12 structure.
  # @param data [String]
  # @return [Hash]
  def cfexaveragestemplateparams(data)
    i_param, _ = data.unpack('vC*')
    i_param_d = {
      0x0000 => :not_offset, # The threshold is not offset by a multiple of the standard deviation.
      0x0001 => :offset1, # The threshold is offset by 1 standard deviation.
      0x0002 => :offset2, # The threshold is offset by 2 standard deviations.
    }[i_param]

    {
      _type: :CFExAveragesTemplateParams,
      iParam: i_param, # iParam (2 bytes): An unsigned integer that specifies the number of standard deviations above or below the average for the rule.
      iParam_d: i_param_d,
      # reserved (14 bytes): MUST be zero and MUST be ignored.
    }
  end

  # 2.5.24 CFExDateTemplateParams
  # The CFExDateTemplateParams structure specifies parameters for the date-related conditional formatting rules specified by a CF12 record or CFExNonCF12 structure.
  # @param data [String]
  # @return [Hash]
  def cfexdatetemplateparams(data)
    date_op, _ = data.unpack('vC*')

    {
      _type: :CFExDateTemplateParams,
      dateOp: date_op # dateOp (2 bytes): An unsigned integer that specifies the type of date comparison.
    }
  end

  # 2.5.25 CFExDefaultTemplateParams
  # This structure specifies that there are no parameters for extensions to conditional formatting rules specified by CFEx.
  # @param _ [String]
  # @return [Hash]
  def cfexdefaulttemplateparams(_)
    {
      _type: :CFExDefaultTemplateParams,
    }
  end

  # 2.5.26 CFExFilterParams
  # @param data [String]
  # @return [Hash]
  def cfexfilterparams(data)
    a_r1_raw, i_param, _ = data.unpack('CvC*')
    a_r1 = Unxls::BitOps.new(a_r1_raw)

    {
      _type: :CFExFilterParams,
      fTop: a_r1.set_at?(0), # A - fTop (1 bit): A bit that specifies whether the top or bottom items are displayed with the conditional formatting.
      fPercent: a_r1.set_at?(1), # B - fPercent (1 bit): A bit that specifies whether a percentage of the top or bottom items are displayed with the conditional formatting, or whether a set number of the top or bottom items are displayed with the conditional formatting.
      # reserved1 (6 bits): MUST be zero and MUST be ignored.
      iParam: i_param, # iParam (2 bytes): An unsigned integer that specifies how many values are displayed with the conditional formatting. If fPercent equals 1 then this field represents a percent and MUST be less than or equal to 100. Otherwise, this field represents a set number of cells and MUST be less than or equal to 1000.
      # reserved2 (13 bytes): MUST be zero and MUST be ignored.
    }
  end

  # 2.5.28 CFExTemplateParams
  # @param data [String]
  # @param icf_template [Integer]
  # @return [Hash]
  def cfextemplateparams(data, icf_template)
    _method = case icf_template
    when 0x05 then :cfexfilterparams
    when 0x08 then :cfextexttemplateparams
    when 0x0F..0x18 then :cfexdatetemplateparams
    when 0x19, 0x1A, 0x1D, 0x1E then :cfexaveragestemplateparams
    else :cfexdefaulttemplateparams
    end

    self.send(_method, data)
  end

  # 2.5.29 CFExTextTemplateParams
  # The CFExTextTemplateParams structure specifies parameters for text-related conditional formatting rules as specified by a CF12 record or CFExNonCF12 structure.
  # @param data [String]
  # @return [Hash]
  def cfextexttemplateparams(data)
    ctp, _ = data.unpack('vC*')
    ctp_d = {
      0x0000 => :'Text contains',
      0x0001 => :'Text does not contain',
      0x0002 => :'Text begins with',
      0x0003 => :'Text ends with',
    }[ctp]

    {
      _type: :CFExTextTemplateParams,
      ctp: ctp, # ctp (2 bytes): An unsigned integer that specifies the type of text rule.
      ctp_d: ctp_d,
      # reserved (14 bytes): MUST be zero and MUST be ignored.
    }
  end

  # 2.5.30 CFFilter
  # The CFFilter structure specifies the parameters of a conditional formatting rule of type top N filter.
  # @param io [StringIO]
  # @return [Hash]
  def cffilter(io)
    cb_filter, _, a_r2_raw, i_param = io.read(6).unpack('vCCv')
    a_r2 = Unxls::BitOps.new(a_r2_raw)

    {
      _type: :CFFilter,
      cbFilter: cb_filter, # cbFilter (2 bytes): An unsigned integer that specifies the size of the structure in bytes, excluding the cbFilter field itself.
      fTop: a_r2.set_at?(0), # A - fTop (1 bit): A bit that specifies whether the top or bottom items are displayed with the conditional formatting.
      fPercent: a_r2.set_at?(1), # B - fPercent (1 bit): A bit that specifies whether a percentage of top or bottom items are displayed with the conditional formatting, or a set number of top or bottom items are displayed with the conditional formatting.
      # reserved2 (6 bits): MUST be zero and MUST be ignored.
      iParam: i_param # iParam (2 bytes): An unsigned integer that specifies how many values are displayed with the conditional formatting. If fPercent is set to 1 then this field represents a percent and MUST be less than or equal to 100, otherwise this field is a number of cells and MUST be less than or equal to 1000.
    }
  end

  # 2.5.32 CFGradient
  # The CFGradient structure specifies the parameters of a conditional formatting rule that uses color scale formatting.
  # @param io [StringIO]
  # @return [Hash]
  def cfgradient(io)
    _, _, c_interp_curve, c_gradient_curve, attrs = io.read(6).unpack('vC4')
    # unused (2 bytes): Undefined and MUST be ignored.
    # reserved1 (1 byte): MUST be zero and MUST be ignored.
    attrs = Unxls::BitOps.new(attrs)

    result = {
      _type: :CFGradient,
      cInterpCurve: c_interp_curve, # cInterpCurve (1 byte): An unsigned integer that specifies the number of control points in the interpolation curve. It MUST be 0x2 or 0x3.
      cGradientCurve: c_gradient_curve, # cGradientCurve (1 byte): An unsigned integer that specifies the number of control points in the gradient curve. It MUST be equal to cInterpCurve.
      fClamp: attrs.set_at?(0), # A - fClamp (1 bit): A bit that specifies that the cell values are not used when they are out of the range of the interpolation curve. The minimum or the maximum of the interpolation curve is used instead of the cell value.
      fBackground: attrs.set_at?(1), # B - fBackground (1 bit): A bit that specifies that the color scale formatting applies to the background of the cells. It MUST be 1.
      # reserved2 (6 bits): MUST be zero and MUST be ignored.
    }

    result[:rgInterp] = c_interp_curve.times.map { cfgradientinterpitem(io) } # rgInterp (variable): An array of CFGradientInterpItem. Each element is a control point of the interpolation curve. Its element count MUST be cInterpCurve.
    result[:rgCurve] = c_gradient_curve.times.map { cfgradientitem(io) } # rgCurve (variable): An array of CFGradientItem. Each element is a control point of the gradient curve. Its element count MUST be cGradientCurve.

    result
  end

  # 2.5.33 CFGradientInterpItem
  # The CFGradientInterpItem structure specifies one control point in the interpolation curve.
  # @param io [StringIO]
  # @return [Hash]
  def cfgradientinterpitem(io)
    {
      cfvoInterp: cfvo(io), # cfvoInterp (variable): A CFVO structure that specifies the cell value associated with the numerical value specified in numDomain.
      numDomain: xnum(io.read(8)) # numDomain (8 bytes): An Xnum (section 2.5.342) structure that specifies the numerical value of this control point.
    }
  end

  # 2.5.34 CFGradientItem
  # The CFGradientItem structure specifies one control point in the gradient curve. The gradient curve specifies a color scale used in conditional formatting and maps numerical values to colors.
  # @param io [StringIO]
  # @return [Hash]
  def cfgradientitem(io)
    {
      numGrange: xnum(io.read(8)), # numGrange (8 bytes): An Xnum (section 2.5.342) that specifies the numerical value of the control point.
      color: cfcolor(io) # A CFColor that specifies the color associated with the numerical value specified in numGrange.
    }
  end

  # 2.5.35 CFMStateItem
  # The CFMStateItem structure specifies the threshold value associated with an icon for a CFMultistate conditional formatting rule.
  # @param io [StringIO]
  # @return [Hash]
  def cfmstateitem(io)
    result = {}

    result[:cfvo] = cfvo(io) # cfvo (variable): A CFVO that specifies the threshold value.
    result[:fEqual] = io.read(1).unpack('C').first == 0x01 # fEqual (1 byte): Cell values that are equal to the threshold value pass the threshold.

    io.read(4) # unused (4 bytes): Undefined and MUST be ignored.

    result
  end

  # 2.5.36 CFMultistate
  # The CFMultistate structure specifies the parameters for a conditional formatting rule that represents cell values with icons from an icon set.
  # @param io [StringIO]
  # @return [Hash]
  def cfmultistate(io)
    _, _, c_states, i_icon_set, a_r3_raw = io.read(6).unpack('vCCCC')
    a_r3 = Unxls::BitOps.new(a_r3_raw)

    result = {
      _type: :CFMultistate,
      # unused (2 bytes): Undefined and MUST be ignored.
      # reserved1 (1 byte): MUST be zero and MUST be ignored.
      cStates: c_states, # cStates (1 byte): An unsigned integer that specifies the number of items in the icon set.
      iIconSet: i_icon_set, # iIconSet (1 byte): An unsigned integer that specifies the icon set that represents the cell values.
      fIconOnly: a_r3.set_at?(0), # A - fIconOnly (1 bit): A bit that specifies whether only the icon will be displayed in the sheet and that the cell value will be hidden.
      # B - reserved2 (1 bit): MUST be zero and MUST be ignored.
      fReverse: a_r3.set_at?(2), # C - fReverse (1 bit): A bit that specifies whether the order of the icons in the set is reversed.
      # reserved3 (5 bits): MUST be zero and MUST be ignored.
    }

    result[:rgStates] = c_states.times.map { cfmstateitem(io) } # rgStates (variable): An array of CFMStateItem. Each element specifies a threshold for the respective icon in the set, below which cell values are represented by the next icon in the set. The element count MUST be equal to cStates.

    result
  end

  # 2.5.198.5 CFParsedFormula
  # The CFParsedFormula structure specifies a formula (section 2.2.2) used in a conditional formatting rule.
  # @param io [StringIO]
  # @return [Hash]
  def cfparsedformula(io)
    cce = io.read(2).unpack('v').first
    io.read(cce)

    {
      cce: cce, # cce (2 bytes): An unsigned integer that specifies the length of rgce in bytes.
      rgce: :not_implemented # rgce (variable): An Rgce that specifies the sequence of Ptg structures for the formula.
    }
  end

  # 2.5.198.6 CFParsedFormulaNoCCE
  # The CFParsedFormulaNoCCE structure specifies a formula (section 2.2.2) used in a conditional formatting rule, in a CF or CF12 record in which the size of the formula in bytes is specified.
  # @param _ [String]
  # @return [Hash]
  def cfparsedformulanocce(_)
    {
      rgce: :not_implemented
    }
  end

  # 2.5.39 CFVO
  # The CFVO structure specifies a Conditional Formatting Value Object (CFVO) that specifies how to calculate a value from the range of cells that a conditional formatting rule applies to.
  # @param io [StringIO]
  # @return [Hash]
  def cfvo(io)
    cfvo_type = io.read(1).unpack('C').first
    cfvo_type_d = {
      0x01 => :value, # X
      0x02 => :range_min, # The minimum value from the range of cells that the conditional formatting rule applies to.
      0x03 => :range_max, # The maximum value from the range of cells that the conditional formatting rule applies to.
      0x04 => :range_min_x_percentile, # The minimum value in the range of cells that the conditional formatting rule applies to plus X percent of the difference between the maximum and minimum values in the range of cells that the conditional formatting rule applies to. For example, if the min and max values in the range are 1 and 10 respectively, and X is 10, then the CFVO value is 1.9.
      0x05 => :range_max_x_percentile, # The minimum value of the cell that is in X percentile of the range of cells that the conditional formatting rule applies to.
      0x07 => :formula_result  # The result of evaluating fmla.
    }[cfvo_type]

    result = {
      cfvoType: cfvo_type, # cfvoType (1 byte): An unsigned integer that specifies how the CFVO value is determined.
      cfvoType_d: cfvo_type_d,
      fmla: cfvoparsedformula(io), # fmla (variable): A CFVOParsedFormula that specifies the formula used to calculate the CFVO value.
    }

    unless %i(range_min range_max).include?(cfvo_type_d) && result[:fmla][:cce].zero?
      result[:numValue] = xnum(io.read(8)) # numValue (8 bytes): An Xnum (section 2.5.342) that specifies a static value used to calculate the CFVO value.
    end

    result
  end

  # 2.5.198.7 CFVOParsedFormula
  # The CFVOParsedFormula structure specifies a formula (section 2.2.2) without relative references that is used in a conditional formatting rule.
  # @param io [StringIO]
  # @return [Hash]
  alias :cfvoparsedformula :cfparsedformula

  # See 2.4.122 Font, bCharSet
  # @param id [Integer]
  # @return [Symbol]
  def _character_set(id)
    Unxls::Biff8::Constants::CHARSETS[id]
  end

  # 2.5.51 ColRelU
  # @param value [Integer]
  # @return [Hash]
  def colrelu(value)
    attrs = Unxls::BitOps.new(value)
    {
      col: attrs.value_at(0..13), # zero-based index of a column in the sheet
      colRelative: attrs.set_at?(14), # specifies whether col is a relative reference
      rowRelative: attrs.set_at?(15) # specifies whether a row index corresponding to col in the structure containing this structure is a relative reference
    }
  end

  # 2.5.56 CondFmtStructure, page 621
  # @param io [StringIO]
  # @return [Hash]
  def condfmtstructure(io)
    ccf, a_nid_raw = io.read(4).unpack('vv')
    a_nid = Unxls::BitOps.new(a_nid_raw)
    result = {
      ccf: ccf, # ccf (2 bytes): An unsigned integer that specifies the count of CF12 records that follow the containing record.
      fToughRecalc: a_nid.set_at?(0), # A - fToughRecalc (1 bit): A bit that specifies that the appearance of the cell requires significant processing.
      nID: a_nid.value_at(1..15) # nID (15 bits): An unsigned integer that identifies this record.
    }

    result[:refBound] = ref8u(io.read(8)) # refBound (8 bytes): A Ref8U structure that specifies bounds of the set of cells to which the rules are applied.
    result[:sqref] = sqrefu(io) # sqref (variable): A SqRefU structure that specifies the cells to which the conditional formatting rules apply.

    result
  end

  # 2.5.91 DXFALC, page 638
  # @param io [StringIO]
  # @return [Hash]
  def dxfalc(io)
    result = _xfalc(io) # bytes 1-3
    %i(fAtrNum, fAtrFnt, fAtrAlc, fAtrBdr, fAtrPat, fAtrProt).each do |k|
      result.delete(k) # 24…31 unused (8 bits): Undefined and MUST be ignored.
    end
    result[:iIndent] = io.read(4).unpack('l<').first # A signed integer that specifies the relative level of indentation. The relative level of indentation will be added to any previous indentation.
    result
  end

  # 2.5.92 DXFBdr, page 639
  # @param io [StringIO]
  # @return [Hash]
  def dxfbdr(io)
    result = _xfbdr(io) # 8 bytes
    %i(fHasXFExt fls fls_d).each do |k|
      result.delete(k) # 25…32 unused (7 bits): Undefined and MUST be ignored.
    end
    result
  end

  # 2.5.93 DXFFntD, page 640
  # Specifies a font and its format attributes.
  # @param io [StringIO]
  # @return [Integer]
  def dxffntd(io)
    result = {
      cchFont: (cch_font = io.read(1).unpack('C').first) # cchFont (1 byte): An unsigned integer that specifies the number of characters of the font name string.
    }

    st_font_name_raw = io.read(63)
    if cch_font > 0
      st_font_name_io = StringIO.new(st_font_name_raw)
      result[:stFontName] = xlunicodestringnocch(st_font_name_io, cch_font) # stFontName (variable): An XLUnicodeStringNoCch that specifies the font name. MUST exist if and only if cchFont is greater than zero. The number of characters in the string is specified in cchFont.
    end

    result[:stxp] = stxp(io.read(16)) # stxp (16 bytes): A Stxp that specifies the font attributes.

    result[:icvFore] = io.read(4).unpack('l<').first # icvFore (4 bytes): An integer that specifies the color of the font. The value MUST be -1, 32767 or any of the valid values of the IcvFont structure. A value of -1 specifies that this value is ignored. A value of 32767 specifies that the color of the font is the default foreground text color. Any other value specifies the color of the font as specified in the IcvFont structure.
    io.read(4) # reserved (4 bytes): MUST be zero, and MUST be ignored.
    result[:tsNinch] = ts(io.read(4).unpack('V').first) # tsNinch (4 bytes): A Ts structure that specifies how the value of stxp.ts is to be interpreted. If tsNinch.ftsItalic is set to 1 then the value of stxp.ts.ftsItalic MUST be ignored. If tsNinch.ftsStrikeout is set to 1 then the value of the stxp.ts.ftsStrikeout MUST be ignored.

    f_sss_n_inch, f_uls_n_inch, f_bls_n_inch, _, ich, cch, i_fnt = io.read(26).unpack('VVVVl<l<v')

    result[:fSssNinch] = f_sss_n_inch == 1 # fSssNinch (4 bytes): A Boolean (section 2.5.14) that specifies whether the value of stxp.sss MUST be ignored.
    result[:fUlsNinch] = f_uls_n_inch == 1 # fUlsNinch (4 bytes): A Boolean that specifies whether the value of stxp.uls MUST be ignored.
    result[:fBlsNinch] = f_bls_n_inch == 1 # fBlsNinch (4 bytes): A Boolean that specifies whether the value of stxp.bls MUST be ignored.
    # unused2 (4 bytes): Undefined and MUST be ignored.
    result[:ich] = ich # ich (4 bytes): A signed integer that specifies the zero based index of the first character to which this font applies. MUST be greater than or equal to 0xFFFFFFFF. MUST be set to 0xFFFFFFFF when the font is to be updated.
    result[:cch] = cch # cch (4 bytes): A signed integer that specifies the number of characters to which this font applies. MUST be greater than or equal to ich field. MUST be set to 0xFFFFFFFF if the ich field is set to 0xFFFFFFFF.
    result[:iFnt] = i_fnt # iFnt (2 bytes): An unsigned integer that specifies the font. If the value is 0 then the default font is used. If the value is greater than 0 then the font to be applied is determined by the font name specified in stFontName.

    result
  end

  # 2.5.95 DXFN, page 641
  # The DXFN structure specifies differential formatting.
  # @param io [StringIO]
  # @return [Hash]
  def dxfn(io)
    a_p_raw, q_d_raw, e_h_raw = io.read(6).unpack('v3')

    a_p = Unxls::BitOps.new(a_p_raw)
    result = {
      alchNinch: a_p.set_at?(0), # A - alchNinch (1 bit): A bit that specifies whether the value of dxfalc.alc MUST be ignored.
      alcvNinch: a_p.set_at?(1), # B - alcvNinch (1 bit): A bit that specifies whether the value of dxfalc.alcv MUST be ignored.
      wrapNinch: a_p.set_at?(2), # C - wrapNinch (1 bit): A bit that specifies whether the value of dxfalc.fWrap MUST be ignored.
      trotNinch: a_p.set_at?(3), # D - trotNinch (1 bit): A bit that specifies whether the value of dxfalc.trot MUST be ignored.
      kintoNinch: a_p.set_at?(4), # E - kintoNinch (1 bit): A bit that specifies whether the value of dxfalc.fJustLast MUST be ignored.
      cIndentNinch: a_p.set_at?(5), # F - cIndentNinch (1 bit): A bit that specifies whether the values of dxfalc.cIndent and dxfalc.iIndent MUST be ignored.
      fShrinkNinch: a_p.set_at?(6), # G - fShrinkNinch (1 bit): A bit that specifies whether the value of dxfalc.fShrinkToFit MUST be ignored.
      fMergeCellNinch: a_p.set_at?(7), # H - fMergeCellNinch (1 bit): A bit that specifies whether the value of dxfalc.fMergeCell MUST be ignored.
      lockedNinch: a_p.set_at?(8), # I - lockedNinch (1 bit): A bit that specifies whether the value of dxfprot.fLocked MUST be ignored.
      hiddenNinch: a_p.set_at?(9), # J - hiddenNinch (1 bit): A bit that specifies whether the value of dxfprot.fHidden MUST be ignored.
      glLeftNinch: a_p.set_at?(10), # K - glLeftNinch (1 bit): A bit that specifies whether the values of dxfbdr.dgLeft and dxfbdr.icvLeft MUST be ignored.
      glRightNinch: a_p.set_at?(11), # L - glRightNinch (1 bit): A bit that specifies whether the values of dxfbdr.dgRight and dxfbdr.icvRight MUST be ignored.
      glTopNinch: a_p.set_at?(12), # M - glTopNinch (1 bit): A bit that specifies whether the values of dxfbdr.dgTop and dxfbdr.icvTop MUST be ignored.
      glBottomNinch: a_p.set_at?(13), # N - glBottomNinch (1 bit): A bit that specifies whether the values of dxfbdr.dgBottom and dxfbdr.icvBottom MUST be ignored.
      glDiagDownNinch: a_p.set_at?(14), # O - glDiagDownNinch (1 bit): A bit that specifies whether the value of dxfbdr.bitDiagDown MUST be ignored. When both glDiagDownNinch and glDiagUpNinch are set to 1, the values of dxfbdr.dgDiag and dxfbdr.icvDiag MUST be ignored.
      glDiagUpNinch: a_p.set_at?(15), # P - glDiagUpNinch (1 bit): A bit that specifies whether the value of dxfbdr.bitDiagUp MUST be ignored. When both glDiagDownNinch and glDiagUpNinch are set to 1, the values of dxfbdr.dgDiag and dxfbdr.icvDiag MUST be ignored.
    }

    q_d = Unxls::BitOps.new(q_d_raw)
    result.merge!({
      flsNinch: q_d.set_at?(0), # Q - flsNinch (1 bit): A bit that specifies whether the value of dxfpat.fls MUST be ignored.
      icvFNinch: q_d.set_at?(1), # R - icvFNinch (1 bit): A bit that specifies whether the value of dxfpat.icvForeground MUST be ignored.
      icvBNinch: q_d.set_at?(2), # S - icvBNinch (1 bit): A bit that specifies whether the value of dxfpat.icvBackground MUST be ignored.
      ifmtNinch: q_d.set_at?(3), # T - ifmtNinch (1 bit): A bit that specifies whether the value of dxfnum.ifmt MUST be ignored.
      fIfntNinch: q_d.set_at?(4), # U - fIfntNinch (1 bit): A bit that specifies whether the value of dxffntd.ifnt MUST be ignored.
      # (5) V - unused1 (1 bit): Undefined and MUST be ignored.
      # (6…8) W - reserved1 (3 bits): MUST be zero and MUST be ignored.
      ibitAtrNum: q_d.set_at?(9), # X - ibitAtrNum (1 bit): A bit that specifies whether number formatting information is part of this structure.
      ibitAtrFnt: q_d.set_at?(10), # Y - ibitAtrFnt (1 bit): A bit that specifies whether font information is part of this structure.
      ibitAtrAlc: q_d.set_at?(11), # Z - ibitAtrAlc (1 bit): A bit that specifies whether alignment information is part of this structure.
      ibitAtrBdr: q_d.set_at?(12), # a - ibitAtrBdr (1 bit): A bit that specifies whether border formatting information is part of this structure.
      ibitAtrPat: q_d.set_at?(13), # b - ibitAtrPat (1 bit): A bit that specifies whether pattern information is part of this structure.
      ibitAtrProt: q_d.set_at?(14), # c - ibitAtrProt (1 bit): A bit that specifies whether rotation information is part of this structure.
      iReadingOrderNinch: q_d.set_at?(15), # d - iReadingOrderNinch (1 bit): A bit that specifies whether the value of dxfalc.iReadingOrder MUST be ignored.
    })

    e_h = Unxls::BitOps.new(e_h_raw)
    result.merge!({
      fIfmtUser: e_h.set_at?(0), # e - fIfmtUser (1 bit): A bit that specifies that the number format used is a user-defined format string. When set to 1, dxfnum contains a format string.
      # (1) f - unused2 (1 bit): Undefined and MUST be ignored.
      fNewBorder: (f_new_border = e_h.value_at(2)), # g - fNewBorder (1 bit): A bit that specifies how the border formats apply to a range of cells.
      fNewBorder_d: {
        0 => :all_cells, # Border formats apply to all cells in the range.
        1 => :outline # Border formats only apply to the outline of the range.
      }[f_new_border],
      # (3…14) reserved2 (12 bits): MUST be zero and MUST be ignored.
      fZeroInited: e_h.set_at?(15) # h - fZeroInited (1 bit): A bit that specifies whether the value of dxfalc.iReadingOrder MUST be taken into account.
    })

    if result[:ibitAtrNum]
      # 2.5.99 DXFNum
      result[:dxfnum] = result[:fIfmtUser] ? dxfnumusr(io) : dxfnumifmt(io) # dxfnum (variable): A DXFNum that specifies the number formatting. MUST exist if and only if ibitAtrNum is nonzero.
    end

    if result[:ibitAtrFnt]
      result[:dxffntd] = dxffntd(io) # dxffntd (variable): A DXFFntD that specifies the font. MUST exist if and only if ibitAtrFnt is nonzero.
    end

    if result[:ibitAtrAlc]
      result[:dxfalc] = dxfalc(io) # (8 bytes): A DXFALC that specifies the text alignment properties.
    end

    if result[:ibitAtrBdr]
      result[:dxfbdr] = dxfbdr(io) # (8 bytes): A DXFBdr that specifies the border properties.
    end

    if result[:ibitAtrPat]
      result[:dxfpat] = dxfpat(io) # (4 bytes): A DXFPat that specifies the pattern and colors.
    end

    if result[:ibitAtrProt]
      result[:dxfprot] = dxfprot(io) # (2 bytes): A DXFProt that specifies the protection attributes.
    end

    result
  end

  # 2.5.96 DXFN12
  # The DXFN12 structure specifies differential formatting and is an extension to DXFN.
  # @param io [StringIO]
  # @return [Integer]
  def dxfn12(io)
    cb_dxf = io.read(4).unpack('V').first
    result = {
      cbDxf: cb_dxf # cbDxf (4 bytes): An unsigned integer that specifies the size of the structure in bytes. If greater than zero, it MUST be the total byte count of dfxn and xfext. Otherwise it MUST be zero.
    }

    if cb_dxf.zero?
      io.read(2) # reserved (2 bytes): MUST be zero and MUST be ignored. MUST be omitted when cbDxf is greater than zero.
    else
      pos = io.pos
      result[:dxfn] = dxfn(io) # dxfn (variable): A DXFN that specifies part of the differential formatting. MUST be omitted if cbDxf is 0x00000000.
      dxfn_size = io.pos - pos
      result[:xfext] = xfextnofrt(io) if cb_dxf > dxfn_size # xfext (variable): An XFExtNoFRT that specifies extensions for the differential formatting. MUST be omitted if cbDxf is equal to the byte count of dxfn.
    end

    result
  end

  # 2.5.97 DXFN12List
  # The DXFN12List structure specifies differential formatting used by table block-level formatting. This structure also specifies extensions to the DXFN formatting properties.
  # @param data [String, StringIO]
  # @return [Integer]
  def dxfn12list(data)
    io = data.to_sio

    result = {
      dxfn: dxfn(io) # dxfn (variable): A DXFN structure that specifies differential formatting used by table block-level formatting.
    }

    unless io.eof?
      result[:xfext] = xfextnofrt(io) # xfext (variable): An XFExtNoFRT structure that specifies the set of extensions to the differential formatting properties specified in dxfn. MUST exist if and only if the size of this structure is greater than the size of the dxfn field.
    end

    result
  end

  # 2.5.98 DXFN12NoCB
  # @param io [StringIO]
  # @return [Integer]
  alias :dxfn12nocb :dxfn12list

  # 2.5.100 DXFNumIFmt
  # @param io [StringIO]
  # @return [Integer]
  def dxfnumifmt(io)
    # (0…7) unused (8 bits): Undefined and MUST be ignored.
    # (8…16) Specifies the identifier of a number format.
    io.read(2).unpack('CC').last # See 2.5.165 IFmt
  end

  # 2.5.101 DXFNumUsr
  # @param io [StringIO]
  # @return [String]
  def dxfnumusr(io)
    cb = io.read(2).unpack('v').first

    {
      cb: cb, # Specifies the size of this structure, in bytes.
      fmt: xlunicodestring(io) # Specifies the number format to use as specified in the stFormat field of Format.
    }
  end

  # 2.5.102 DXFPat, page 646
  # @param io [StringIO]
  # @return [Hash]
  def dxfpat(io)
    u_a_raw = io.read(4).unpack('V').first
    u_a = Unxls::BitOps.new(u_a_raw)

    {
      # 0…9 unused1 (10 bits): Undefined and MUST be ignored.
      fls: (fls = u_a.value_at(10..15)), # fls (6 bits): A FillPattern that specifies the fill pattern.
      fls_d: fillpattern(fls),
      icvFore: u_a.value_at(16..22), # @todo _icv_d icvForeground (7 bits): An unsigned integer that specifies the color of the foreground of the cell. The value MUST be an IcvXF value. This value is unused and MUST be ignored if the icvFNinched field in the containing DXFN structure is 1.
      icvBack: u_a.value_at(23..29), # icvBackground (7 bits): An unsigned integer that specifies the color of the background of the cell. The value MUST be an IcvXF value. This value is unused and MUST be ignored if the icvBNinched field in the containing DXFN structure is 1.
      # 30…31 A - unused2 (2 bits): Undefined and MUST be ignored.
    }
  end

  # 2.5.103 DXFProt, page 647
  # @param io [StringIO]
  # @return [Hash]
  def dxfprot(io)
    a_b_raw = io.read(2).unpack('v').first
    a_b = Unxls::BitOps.new(a_b_raw)

    {
      fLocked: a_b.set_at?(0), # A - fLocked (1 bit): A bit that specifies if the cell content is locked when the workbook is protected.
      fHidden: a_b.set_at?(1), # B - fHidden (1 bit): A bit that specifies if the cell content is hidden when the workbook is protected.
      # reserved (14 bits): MUST be zero and MUST be ignored.
    }
  end

  # @param string [String]
  # @param encoding [String]
  # @return [String]
  def _encode_string(string, encoding = Encoding::UTF_16LE)
    string.force_encoding(encoding).encode(Encoding::UTF_8)
  end

  # @param high_byte [TrueClass, FalseClass]
  # @return [Array]
  def _encoding_params(high_byte)
    high_byte ? [Encoding::UTF_16LE, 2] : [Encoding::UTF_8, 1]
  end

  # 2.5.108 ExtProp
  # The ExtProp structure specifies an extension to a formatting property.
  # @param io [StringIO]
  # @return [Hash]
  def extprop(io)
    ext_type, cb = io.read(4).unpack('vv')
    prop_data = _ext_prop_d(ext_type)

    result = {
      extType: ext_type, # extType (2 bytes): An unsigned integer that specifies the type of the extension.
      extType_d: prop_data[:structure],
      _property: prop_data[:property],
      cb: cb, # cb (2 bytes): An unsigned integer that specifies the size of this ExtProp structure.
    }

    ext_prop_data = io.read(cb - 4) # structure size minus ext_type and cb
    parser_method = prop_data[:structure].downcase
    result[:extPropData] = self.send(parser_method, ext_prop_data) # extPropData (variable): This field specifies the extension data.

    result
  end

  # 2.5.109 ExtRst
  # The ExtRst structure specifies phonetic string data.
  # @param data [String]
  # @return [Hash]
  def extrst(data)
    io = StringIO.new(data)

    _, cb = io.read(4).unpack('vv')
    # reserved (2 bytes): MUST be 1, and MUST be ignored.
    result = { cb: cb } # cb (2 bytes): An unsigned integer that specifies the size, in bytes, of the phonetic string data.

    result[:phs] = phs(io.read(4)) # phs (4 bytes): A Phs that specifies the formatting information for the phonetic string.
    result[:rphssub] = rphssub(io) # rphssub (variable): An RPHSSub that specifies the phonetic string.

    result[:rphssub][:crun].times do # See 2.5.219 RPHSSub
      result[:rgphruns] ||= []
      result[:rgphruns] << rgphruns(io)
    end

    result
  end

  # See 2.5.108 ExtProp, p. 649
  # @param id [Integer]
  # @return [Hash]
  def _ext_prop_d(id)
    {
      0x0004 => { structure: :FullColorExt,  property: :'cell interior foreground color' },
      0x0005 => { structure: :FullColorExt,  property: :'cell interior background color' },
      0x0006 => { structure: :XFExtGradient, property: :'cell interior gradient fill' },
      0x0007 => { structure: :FullColorExt,  property: :'top cell border color' },
      0x0008 => { structure: :FullColorExt,  property: :'bottom cell border color' },
      0x0009 => { structure: :FullColorExt,  property: :'left cell border color' },
      0x000A => { structure: :FullColorExt,  property: :'right cell border color' },
      0x000B => { structure: :FullColorExt,  property: :'diagonal cell border color' },
      0x000D => { structure: :FullColorExt,  property: :'cell text color' },
      0x000E => { structure: :FontScheme,    property: :'font scheme' },
      0x000F => { structure: :_int1b,        property: :'text indentation level' }
    }[id]
  end

  # @todo 2.5.112 Feat11FdaAutoFilter
  # The Feat11FdaAutoFilter structure specifies the definition of an automatically generated filter, or AutoFilter.
  # @param io [StringIO]
  # @return [Symbol]
  def feat11fdaautofilter(io)
    cb_auto_filter, _ = io.read(6).unpack('Vv')
    io.pos += cb_auto_filter

    :not_implemented
  end

  # @todo 2.5.113 Feat11FieldDataItem
  # The Feat11FieldDataItem structure specifies a column of a table.
  # @param io [StringIO]
  # @param tft [Hash] Parent TableFeatureType structure
  # @return [Hash]
  def feat11fielddataitem(io, tft)
    id_field, lfdt, lfxidt, ilta, cb_fmt_agg, istn_agg = io.read(24).unpack('V6')

    result = {
      idField: id_field, # idField (4 bytes): An unsigned integer that specifies the identifier of the column. MUST be nonzero and MUST be unique within the FieldData array in the containing TableFeatureType structure.
      lfdt: lfdt, # lfdt (4 bytes): An unsigned integer that specifies the column’s Web based data provider data type. If the lt field of the containing TableFeatureType structure is not set to 0x00000001, this field MUST be 0x00000000; otherwise it MUST be a value from the following table. For more information about the data types, see [MS-WSSTS] section 2.3.
      lfxidt: lfxidt, # lfxidt (4 bytes): An unsigned integer that specifies the column’s XML data type. If the lt field of the containing TableFeatureType structure is not set to 0x00000002, this field MUST be 0x00000000; otherwise it MUST be a value from the following table. For more information about the data types, see [MSDN-SOM].
      ilta: ilta, # ilta (4 bytes): An unsigned integer that specifies the aggregation function to use for the total row of the column.
      cbFmtAgg: cb_fmt_agg, # cbFmtAgg (4 bytes): An unsigned integer that specifies the size, in bytes, of the dxfFmtAgg field.
      istnAgg: istn_agg, # istnAgg (4 bytes): An unsigned integer that specifies the zero-based index of the Style record in the Globals Substream ABNF that is used for the total row of the column. If this value equals 0xFFFFFFFF, the total row of the column uses built-in table styles.
    }

    result[:lfdt_d] = {
      0x01 => :'Text',
      0x02 => :'Number',
      0x03 => :'Boolean',
      0x04 => :'Date Time',
      0x05 => :'Note',
      0x06 => :'Currency',
      0x07 => :'Lookup',
      0x08 => :'Choice',
      0x09 => :'URL',
      0x0A => :'Counter',
      0x0B => :'Multiple Choices',
    }[lfdt]

    result[:lfxidt_d] = _lfxidt_d(lfxidt)

    result[:ilta_d] = {
      0x00 => :'No formula',
      0x01 => :Average,
      0x02 => :Count,
      0x03 => :'Count numbers',
      0x04 => :Max,
      0x05 => :Min,
      0x06 => :Sum,
      0x07 => :'Standard deviation',
      0x08 => :Variance,
      0x09 => :'Custom formula',
    }[ilta]

    attrs = Unxls::BitOps.new(io.read(4).unpack('V').first)
    result.merge!({
      fAutoFilter: attrs.set_at?(0), # A - fAutoFilter (1 bit):  A bit that specifies whether the column has an AutoFilter.
      fAutoFilterHidden: attrs.set_at?(1), # B - fAutoFilterHidden (1 bit): A bit that specifies whether the column has an AutoFilters that is not displayed. When this field is set to 1, fAutoFilter MUST be set to 1.
      fLoadXmapi: attrs.set_at?(2), # C - fLoadXmapi (1 bit): A bit that specifies whether the rgXmap field is present. MUST be 0 if the lt field of the containing TableFeatureType structure is not equal to 0x00000002.
      fLoadFmla: attrs.set_at?(3), # D - fLoadFmla (1 bit): A bit that specifies whether the fmla field is present for a table whose data source is a Web based data provider list. MUST be 0 if the lt field of the containing TableFeatureType structure is not equal to 0x00000001.
      # 4…5 E - unused1 (2 bits): Undefined, and MUST be ignored.
      # 6 F - reserved2 (1 bit):  MUST be zero, and MUST be ignored.
      fLoadTotalFmla: attrs.set_at?(7), # G - fLoadTotalFmla (1 bit): A bit that specifies whether the totalFmla field is present. SHOULD<165> be 1 if ilta is 0x00000009, MUST be 0 otherwise.
      fLoadTotalArray: attrs.set_at?(8), # H - fLoadTotalArray (1 bit): A bit that specifies whether the formula specified by totalFmla is an array formula. MUST be 0 when fLoadTotalFmla is 0.
      fSaveStyleName: attrs.set_at?(9), # I - fSaveStyleName (1 bit): A bit that specifies whether the dskHdrCache.strStyleName field is present.
      fLoadTotalStr: attrs.set_at?(10), # J - fLoadTotalStr (1 bit): A bit that specifies whether the strTotal field is present. MUST be 0 when ilta is not 0x00000000.
      fAutoCreateCalcCol: attrs.set_at?(11), # K - fAutoCreateCalcCol (1 bit): A bit that specifies whether the column has a calculated column formula. MUST be 0 if the lt field of the containing TableFeatureType structure is set to 0x00000001.
      # unused2 (20 bits): Undefined, and MUST be ignored.
    })

    cb_fmt_insert_row, istn_insert_row = io.read(8).unpack('VV')
    result.merge!({
      cbFmtInsertRow: cb_fmt_insert_row, # cbFmtInsertRow (4 bytes): An unsigned integer that specifies the size, in bytes, of the dxfFmtInsertRow field.
      istnInsertRow: istn_insert_row, # istnInsertRow (4 bytes):  An unsigned integer that specifies the zero-based index of the Style record in the Globals Substream ABNF that is used for the insert row of the column. If this value equals 0xFFFFFFFF, the insert row of the column uses built-in table styles.
      strFieldName: xlunicodestring(io), # strFieldName (variable): An XLUnicodeString that specifies the name of the column, as provided by the data source.
      strCaption: xlunicodestring(io), # strCaption (variable): An XLUnicodeString that specifies the caption of the column.
    })

    if result[:cbFmtAgg] > 0
      result[:dxfFmtAgg] = dxfn12list(io.read(result[:cbFmtAgg])) # dxfFmtAgg (variable): A DXFN12List that specifies the formatting of the total row of the column, if different from the style specified by istnAgg or built-in table styles. This field is present if and only if the cbFmtAgg field is greater than 0x00000000.
    end

    if result[:cbFmtInsertRow] > 0
      result[:dxfFmtInsertRow] = dxfn12list(io.read(result[:cbFmtInsertRow])) # dxfFmtInsertRow (variable): A DXFN12List that specifies the formatting of the insert row of the column, if different from the style specified by istnInsertRow or built-in table styles. This field is present if and only if the cbFmtInsertRow field is more than 0x00000000.
    end

    if result[:fAutoFilter]
      result[:AutoFilter] = feat11fdaautofilter(io) # AutoFilter (variable): A Feat11FdaAutoFilter that specifies the characteristics of the AutoFilter for the column. This field is present if and only if the fAutoFilter field of the containing TableFeatureType structure is set to 1.
    end

    if result[:fLoadXmapi]
      result[:rgXmap] = feat11xmap(io) # rgXmap (variable): A Feat11XMap structure that specifies the mapping to the column data within an XML data source. This field is present if and only if the fLoadXmapi bit is set to 1.
    end

    if result[:fLoadFmla]
      result[:fmla] = feat11fmla(io) # # fmla (variable): A Feat11Fmla structure that specifies the column formula whose data source is a Web based data provider list. The specified formula applies to every row of the column, except the total row and the header row. This field is present if and only if the fLoadFmla bit is set to 1.
    end

    if result[:fLoadTotalFmla]
      result[:totalFmla] = feat11totalfmla(io, result[:fLoadTotalArray]) # totalFmla (variable): A Feat11TotalFmla structure that specifies the formula to use for the total row of the column. This field is present if and only if the fLoadTotalFmla bit is set to 1.
    end

    if result[:fLoadTotalStr]
      result[:strTotal] = xlunicodestring(io) # strTotal (variable): An XLUnicodeString structure that specifies the text to use for the total row of the column. MUST contain less than or equal to 32767 characters. This field is present if and only if the fLoadTotalStr bit is set to 1.
    end

    if tft[:lt] == 0x01
      result[:wssInfo] = feat11wsslistinfo(io, lfdt) # wssInfo (variable): A Feat11WSSListInfo that specifies the relationship between the column and a Web based data provider list. This field is present if and only if the lt field of the containing TableFeatureType structure is set to 0x00000001.
    end

    if tft[:lt] == 0x03
      result[:qsif] = io.read(4).unpack('V').first # qsif (4 bytes): An unsigned integer that specifies the relationship between the column and its Microsoft Query data source. MUST be equal to the idField field of a Qsif record within the Worksheet Substream. This field is present if and only if the lt field of the containing TableFeatureType structure is set to 0x00000003 (External data source). MUST be greater than zero and MUST be unique within the FieldData array in the containing TableFeatureType structure.
    end

    unless tft[:crwHeader] || tft[:fSingleCell]
      result[:dskHdrCache] = cacheddiskheader(io, result[:fSaveStyleName]) # dskHdrCache (variable): A CachedDiskHeader that specifies the column header formatting information. This field is present if and only if the crwHeader field of the containing TableFeatureType structure is set to 0x0000 and the fSingleCell field of the containing TableFeatureType structure is set to 0.
    end

    result

  rescue StandardError => e # Unfinished listparsedarrayformula in feat11totalfmla might raise in some cases
    { _error: e }
  end

  # @todo 2.5.114 Feat11Fmla
  # The Feat11Fmla structure specifies a formula (section 2.2.2) that is used as a column formula.
  # @param io [StringIO]
  # @return [Symbol]
  def feat11fmla(io)
    cb_fmla = io.read(2).unpack('v').first # cbFmla (2 bytes): An unsigned integer that specifies the size, in bytes, of the rgbFmla field.
    io.pos += cb_fmla # rgbFmla (variable): A ListParsedFormula that specifies the parsed expression of the column formula.

    :not_implemented
  end

  # @todo 2.5.118 Feat11TotalFmla
  # The Feat11TotalFmla structure specifies a formula (section 2.2.2) that can be used as a total row formula.
  # @param io [StringIO]
  # @param f_load_total_array [true, false]
  # @return [Symbol]
  def feat11totalfmla(io, f_load_total_array)
    f_load_total_array ? listparsedarrayformula(io) : listparsedformula(io) # rgbFmlaTotal (variable): A ListParsedFormula or ListParsedArrayFormula that specifies the parsed expression of the total row formula. When the fLoadTotalArray field of the containing Feat11FieldDataItem structure is set to 1, this field is a ListParsedArrayFormula; otherwise, it is a ListParsedFormula.

    :not_implemented
  end

  # @todo 2.5.120 Feat11XMap
  # The Feat11XMap structure specifies the mapping between a table column’s data and an XML data source.
  # @param io [StringIO]
  # @return [Symbol]
  def feat11xmap(io)
    i_xmap_mac = io.read(2).unpack('v').first
    i_xmap_mac.times { feat11xmapentry(io) }

    :not_implemented
  end

  # @todo 2.5.121 Feat11XMapEntry
  # The Feat11XMapEntry structure specifies a mapping to an XML data source.
  # @param io [StringIO]
  # @return [Symbol]
  def feat11xmapentry(io)
    # A - reserved1 (1 bit): MUST be zero, and MUST be ignored.
    # B - fLoadXMap (1 bit): MUST be 1, and MUST be ignored.
    # C - fCanBeSingle (1 bit): A bit that specifies whether details.rgbXPath resolves to a single XML node or a collection of XML nodes.
    # D - reserved2 (1 bit):  MUST be zero, and MUST be ignored.
    # reserved3 (28 bits): MUST be zero, and MUST be ignored.
    io.pos += 4

    feat11xmapentry2(io) # details (variable): A Feat11XMapEntry2 that specifies the mapping between the data and the XML data source.

    :not_implemented
  end

  # @todo 2.5.122 Feat11XMapEntry2
  # The Feat11XMapEntry2 structure specifies the mapping to an XML data source.
  # @param io [StringIO]
  # @return [Symbol]
  def feat11xmapentry2(io)
    io.pos += 4 # dwMapId (4 bytes): An unsigned integer that specifies the XML schema associated with this table column. The value MUST equal the value of the ID attribute of a Map element contained within the XML stream (section 2.1.7.22).
    xlunicodestring(io) # rgbXPath (variable): An XLUnicodeString that contains the XPath expression that specifies the mapped element in the XML schema specified by dwMapId. The length of this string MUST be less than 32000.

    :not_implemented
  end

  # @todo 2.5.119 Feat11WSSListInfo
  # The Feat11WSSListInfo structure specifies the relationship between a table column and a Web-based data provider list.
  # @param io [StringIO]
  # @param lfdt [Integer]
  # @return [Symbol]
  def feat11wsslistinfo(io, lfdt)
    # LCID (4 bytes): An unsigned integer that specifies the language code identifier (LCID) of the source data.
    # cDec (4 bytes): An unsigned integer that specifies the number of decimal places for a numeric column.
    # A - fPercent (1 bit): A bit that specifies whether the numeric values in the column are displayed as percentages.
    # B - fDecSet (1 bit): A bit that specifies whether the numeric values in the column are displayed with a fixed decimal point. The position of the decimal point is specified by the cDec field.
    # C - fDateOnly (1 bit): A bit that specifies whether only the date part of date/time values is displayed.
    # D - fReadingOrder (2 bits): An unsigned integer that specifies the reading order. MUST be a value from the following table:
    # 0x0 Reading order is determined by the application based on the reading order of the cells surrounding the table.
    # 0x1 Reading order is left-to-right.
    # 0x2 Reading order is right-to-left.
    # E - fRichText (1 bit): A bit that specifies whether the column contains rich text.
    # F - fUnkRTFormatting (1 bit): A bit that specifies whether the column contains unrecognized rich text formatting.
    # G - fAlertUnkRTFormatting (1 bit): A bit that specifies whether the column contains unrecognized rich text formatting that requires notifying the user.
    # unused1 (24 bits): Undefined and MUST be ignored.
    io.pos += 12

    # H - fReadOnly (1 bit): A bit that specifies whether the column is read only.
    # I - fRequired (1 bit): A bit that specifies whether every item in this column has to contain data.
    # J - fMinSet (1 bit): A bit that specifies whether a minimum numeric value for the column exists. The minimum value is stored in the List Data stream within the LISTSCHEMA element, under the Field node's Min attribute.
    # K - fMaxSet (1 bit): A bit that specifies whether a maximum numeric value for the column exists. The maximum value is stored in the List Data stream within the LISTSCHEMA element, under the Field node's Max attribute.
    # L - fDefaultSet (1 bit): A bit that specifies whether there is a default value for the column.
    # M - fDefaultDateToday (1 bit): A bit that specifies whether the default value for the column is the current date.
    # N - fLoadFormula (1 bit): A bit that specifies whether a validation formula exists for this column. The formula is specified by the strFormula field.
    # O - fAllowFillIn (1 bit): A bit that specifies whether a choice field allows custom user entries.
    # bDefaultType (8 bits): An unsigned integer that specifies the type of the rgbDV default value. This field MUST be ignored if fDefaultSet is not 0x1; otherwise, it MUST be a value from the following table:
    # 0x00 There is no default value specified.
    # 0x01 rgbDV is a string.
    # 0x02 rgbDV is a Boolean.
    # 0x03 rgbDV is a number.
    # unused2 (16 bits): Undefined, MUST be ignored.
    attrs = Unxls::BitOps.new(io.read(4).unpack('V').first)
    f_default_set = attrs.set_at?(4)
    f_load_formula = attrs.set_at?(6)
    b_default_type = attrs.value_at(8..15)

    # rgbDV (variable): A field of variable data type that specifies the default value for the column. The data type is specified in the lfdt field of the containing Feat11FieldDataItem structure. MUST be one of the data types specified in the following table:
    # 0x01 Short Text # An XLUnicodeString with a maximum length of 255 Unicode characters.
    # 0x08 Choice # An XLUnicodeString with a maximum length of 255 Unicode characters.
    # 0x0B Multi-choice # An XLUnicodeString with a maximum length of 255 Unicode characters.
    # 0x02 Number # An Xnum (section 2.5.342).
    # 0x04 Date time # A DateAsNum. (Xnum)
    # 0x06 Currency # An Xnum.
    # 0x03 Yes/No # A 32-bit Boolean (section 2.5.14).
    # 0x05 Invalid # rgbDV does not exist.
    # 0x07 Invalid # rgbDV does not exist.
    # 0x09 Invalid # rgbDV does not exist.
    # 0x0A Invalid # rgbDV does not exist.
    if f_default_set && !b_default_type.zero?
      case lfdt
      when 0x01, 0x08, 0x0B then xlunicodestring(io)
      when 0x02, 0x04, 0x06 then io.pos += 8
      when 0x03 then io.pos += 4
      else nil
      end
    end

    if f_load_formula
      xlunicodestring(io) # strFormula (variable): An XLUnicodeString that specifies the validation formula as defined by the Web based data provider. This field exists if and only if fLoadFormula is set to 0x1.
    end

    io.pos += 4 # reserved (4 bytes): MUST be 0x00000000, and MUST be ignored.

    :not_implemented
  end

  # 2.5.127 FillPattern
  # @param id [Integer]
  # @return [Symbol]
  def fillpattern(id)
    {
      0x00 => :FLSNULL, # No fill pattern
      0x01 => :FLSSOLID, # Solid
      0x02 => :FLSMEDGRAY, # 50% gray
      0x03 => :FLSDKGRAY, # 75% gray
      0x04 => :FLSLTGRAY, # 25% gray
      0x05 => :FLSDKHOR, # Horizontal stripe
      0x06 => :FLSDKVER, # Vertical stripe
      0x07 => :FLSDKDOWN, # Reverse diagonal stripe
      0x08 => :FLSDKUP, # Diagonal stripe
      0x09 => :FLSDKGRID, # Diagonal crosshatch
      0x0A => :FLSDKTRELLIS, # Thick diagonal crosshatch
      0x0B => :FLSLTHOR, # Thin horizontal stripe
      0x0C => :FLSLTVER, # Thin vertical stripe
      0x0D => :FLSLTDOWN, # Thin reverse diagonal stripe
      0x0E => :FLSLTUP, # Thin diagonal stripe
      0x0F => :FLSLTGRID, # Thin horizontal crosshatch
      0x10 => :FLSLTTRELLIS, # Thin diagonal crosshatch
      0x11 => :FLSGRAY125, # 12.5% gray
      0x12 => :FLSGRAY0625, # 6.25% gray
    }[id]
  end

  # 2.5.133 FormulaValue
  # The FormulaValue structure specifies the current value of a formula. It can be a numeric value, a Boolean value, an error value, a string value, or a blank string value.
  # @param data [String]
  # @return [Hash]
  def formulavalue(data)
    f_expr_o = data[6..7].unpack('v').first

    return { _value: xnum(data), _type: :float } if f_expr_o != 0xFFFF # value is Xnum if last 2 bytes are FFFF

    byte1_val = data[0].unpack('C').first
    value_type = {
      0 => :string,
      1 => :boolean,
      2 => :error,
      3 => :blank_string
    }[byte1_val]

    value = case value_type
    when :string then nil # value is stored in a String record that immediately follows Formula record
    when :boolean then data[2] != "\0" # byte3 specifies Boolean value
    when :error then berr(data[2].unpack('C').first) # byte 3 specifies an error
    when :blank_string then ''
    else raise "Unknown value type #{byte1_val} in FormulaValue structure"
    end

    { _value: value, _type: value_type }
  end

  # See 2.4.122 Font, bFamily
  # @param id [Integer]
  # @return [Symbol]
  def _font_family(id)
    {
      0x00 => :'Not applicable',
      0x01 => :Roman,
      0x02 => :Swiss,
      0x03 => :Modern,
      0x04 => :Script,
      0x05 => :Decorative
    }[id]
  end

  # 2.5.131 FontScheme
  # The FontScheme enumeration specifies the font scheme to which this font belongs.
  # @param data [String]
  # @return [Symbol]
  def fontscheme(data)
    id = data.unpack('C').first

    name = {
      0x00 => :XFSNONE, # No font scheme
      0x01 => :XFSMAJOR, # Major scheme
      0x02 => :XFSMINOR, # Minor scheme
      0xff => :XFSNIL # Ninched state
    }[id]

    {
      FontScheme: id,
      FontScheme_d: name,
    }
  end

  # 2.5.132 FormatRun
  # The FormatRun structure specifies formatting information for a text run.
  # @param io [StringIO]
  # @return [Hash]
  def formatrun(io)
    ich, ifnt = io.read(4).unpack('vv')

    {
      ich: ich, # ich (2 bytes): An unsigned integer that specifies the zero-based index of the first character of the text that contains the text run.
      ifnt: ifnt # ifnt (2 bytes): A FontIndex structure that specifies the font. If ich is equal to the length of the text, this record is undefined and MUST be ignored.
    }
  end

  # 2.5.134 FrtFlags
  # The FrtFlags structure specifies flags used in future record headers.
  # @param data [String]
  # @return [Hash]
  def frtflags(data)
    attrs = Unxls::BitOps.new(data)

    {
      fFrtRef: attrs.set_at?(0), # A - fFrtRef (1 bit): A bit that specifies whether the containing record specifies a range of cells.
      fFrtAlert: attrs.set_at?(1) #  B - fFrtAlert (1 bit): A bit that specifies whether to alert the user of possible problems when saving the file without having recognized this record.
      # reserved (14 bits): MUST be zero, and MUST be ignored.
    }
  end

  # 2.5.135 FrtHeader
  # The FrtHeader structure specifies a future record type header.
  # @param data [String]
  # @return [Hash]
  def frtheader(data)
    rt, grbit_frt = data.unpack('vvC*')

    {
      rt: rt, # rt (2 bytes): An unsigned integer that specifies the record type identifier. MUST be identical to the record type identifier of the containing record.
      grBitFrt: frtflags(grbit_frt) # grbitFrt (2 bytes): An FrtFlags that specifies attributes for this record. The value of grbitFrt.fFrtRef MUST be zero. The value of grbitFrt.fFrtAlert MUST be zero.
      # reserved (8 bytes): MUST be zero, and MUST be ignored.
    }
  end

  # 2.5.137 FrtRefHeader
  # The FrtRefHeader structure specifies a future record type header.
  # @param io [StringIO]
  # @return [Hash]
  def frtrefheader(io)
    result = frtheader(io.read(4)) # rt, grbitFrt
    result[:ref8] = ref8u(io.read(8)) # ref8 (8 bytes): A Ref8U that references the range of cells associated with the containing record.
    result
  end

  # 2.5.139 FrtRefHeaderU, page 681
  # The FrtRefHeaderU structure specifies a future record type header.
  # @param io [StringIO]
  # @return [Hash]
  alias :frtrefheaderu :frtrefheader

  # 2.5.138 FrtRefHeaderNoGrbit
  # @param io [StringIO]
  # @return [Hash]
  def frtrefheadernogrbit(io)
    {
      dt: io.read(2).unpack('v').first,
      ref8: ref8u(io.read(8))
    }
  end

  # 2.5.143 FtCmo
  # The FtCmo structure specifies the common properties of the Obj record that contains this FtCmo.
  # @param data [String]
  def ftcmo(data)
    ft, cb, ot, id, attrs, _ = data.unpack('v5V3')

    ot_d = {
      0x00 => :Group,
      0x01 => :Line,
      0x02 => :Rectangle,
      0x03 => :Oval,
      0x04 => :Arc,
      0x05 => :Chart,
      0x06 => :Text,
      0x07 => :Button,
      0x08 => :Picture,
      0x09 => :Polygon,
      0x0B => :Checkbox,
      0x0C => :'Radio button',
      0x0D => :'Edit box',
      0x0E => :Label,
      0x0F => :'Dialog box',
      0x10 => :'Spin control',
      0x11 => :Scrollbar,
      0x12 => :List,
      0x13 => :'Group box',
      0x14 => :'Dropdown list',
      0x19 => :Note,
      0x1E => :'OfficeArt object',
    }[ot]

    attrs = Unxls::BitOps.new(attrs)

    {
      ft: ft, # ft (2 bytes): Reserved. MUST be 0x15.
      cb: cb, # cb (2 bytes):  Reserved. MUST be 0x12.
      ot: ot, # ot (2 bytes): An unsigned integer that specifies the type of object represented by the Obj record that contains this FtCmo.
      ot_d: ot_d,
      id: id, # id (2 bytes): An unsigned integer that specifies the identifier of this object. This object identifier is used by other types to refer to this object. The value of id MUST be unique among all Obj records within the Chart Sheet Substream ABNF, Macro Sheet Substream ABNF and Worksheet Substream ABNF.
      fLocked: attrs.set_at?(0), # A - fLocked (1 bit): A bit that specifies whether this object is locked.
      # B - reserved (1 bit): Reserved. MUST be 0.
      fDefaultSize: attrs.set_at?(2), # C - fDefaultSize (1 bit): A bit that specifies whether the application is expected to choose the object’s size.
      fPublished: attrs.set_at?(3), # D - fPublished (1 bit): A bit that specifies whether this is a chart object that is expected to be published the next time the sheet containing it is published<172>. This bit is ignored if the fPublishedBookItems field of the BookExt_Conditional12 structure is zero.
      fPrint: attrs.set_at?(4), # E - fPrint (1 bit): A bit that specifies whether the image of this object is intended to be included when printed.
      # F, G - unused1, 2 (1 bit): Undefined and MUST be ignored.
      fDisabled: attrs.set_at?(7), # H - fDisabled (1 bit): A bit that specifies whether this object has been disabled.
      fUIObj: attrs.set_at?(8), # I - fUIObj (1 bit): A bit that specifies whether this is an auxiliary object that can only be automatically inserted by the application (as opposed to an object that can be inserted by a user).
      fRecalcObj: attrs.set_at?(9), # J - fRecalcObj (1 bit): A bit that specifies whether this object is expected to be updated on load to reflect the values in the range associated with the object.
      # K, L - unused3, 4 (1 bit): Undefined and MUST be ignored.
      fRecalcObjAlways: attrs.set_at?(12), # M - fRecalcObjAlways (1 bit): A bit that specifies whether this object is expected to be updated whenever the value of a cell in the range associated with the object changes.
      # N, O, P - unused5, 6, 7 (1 bit): Undefined and MUST be ignored.
      # unused8, 9, 10 (4 bytes):  Undefined and MUST be ignored.
    }
  end

  # 2.5.155 FullColorExt
  # The FullColorExt structure specifies a color.
  # @param data [String]
  # @return [Hash]
  def fullcolorext(data)
    io = StringIO.new(data)
    xclr_type, n_tint_shade = io.read(4).unpack('vs<')

    type = xcolortype(xclr_type)

    {
      xclrType: xclr_type, # xclrType (2 bytes): An XColorType that specifies how the color information is stored.
      xclrType_d: type,
      nTintShade: n_tint_shade, # nTintShade (2 bytes): A signed integer that specifies the tint of the color. Positive values lighten the color, and negative values darken the color.
      nTintShade_d: _map_tint(n_tint_shade),
      xclrValue: _interpret_xclr(type, io.read(4)), # xclrValue (4 bytes): An unsigned integer that specifies the color data.
      # unused (8 bytes): Undefined and MUST be ignored.
    }
  end

  # 2.5.156 GradStop
  # The GradStop structure specifies a gradient stop for a gradient fill.
  # @param io [StringIO]
  # @return [Hash]
  def gradstop(io)
    xclr_type = io.read(2).unpack('v').first
    type = xcolortype(xclr_type)

    result = {
      xclrType: xclr_type, # xclrType (2 bytes): An XColorType that specifies how the color information is stored.
      xclrType_d: type,
      xclrValue: _interpret_xclr(type, io.read(4)), # xclrValue (4 bytes): An unsigned integer that specifies the color data.
      numPosition: xnum(io.read(8)).round(2), # numPosition (8 bytes): An Xnum (section 2.5.342) that specifies the gradient stop position as the percentage of the gradient range. The gradient stop position is the position within the gradient range where this gradient stop’s color begins.
    }

    num_tint = xnum(io.read(8)).round(2)
    result[:numTint] = num_tint # numTint (8 bytes): An Xnum that specifies the tint of the color.
    result[:nTintShade_d] = num_tint # for compability with FullColorExt

    result
  end

  # 2.5.159 HorizAlign
  # @param id [Integer]
  # @return [Symbol]
  def horizalign(id)
    {
      0xFF => :ALCNIL, # Alignment not specified
      0x00 => :ALCGEN, # General alignment
      0x01 => :ALCLEFT, # Left alignment
      0x02 => :ALCCTR, # Centered alignment
      0x03 => :ALCRIGHT, # Right alignment
      0x04 => :ALCFILL, # Fill alignment
      0x05 => :ALCJUST, # Justify alignment
      0x06 => :ALCCONTCTR, # Center-across-selection alignment
      0x07 => :ALCDIST, # Distributed alignment
    }[id]
  end

  def _icf_template_d(value)
    {
      0x0000 => :'Cell value',
      0x0001 => :'Formula',
      0x0002 => :'Color scale formatting',
      0x0003 => :'Data bar formatting',
      0x0004 => :'Icon set formatting',
      0x0005 => :'Filter',
      0x0007 => :'Unique values',
      0x0008 => :'Contains text',
      0x0009 => :'Contains blanks',
      0x000A => :'Contains no blanks',
      0x000B => :'Contains errors',
      0x000C => :'Contains no errors',
      0x000F => :'Today',
      0x0010 => :'Tomorrow',
      0x0011 => :'Yesterday',
      0x0012 => :'Last 7 days',
      0x0013 => :'Last month',
      0x0014 => :'Next month',
      0x0015 => :'This week',
      0x0016 => :'Next week',
      0x0017 => :'Last week',
      0x0018 => :'This month',
      0x0019 => :'Above average',
      0x001A => :'Below average',
      0x001B => :'Duplicate values',
      0x001D => :'Above or equal to average',
      0x001E => :'Below or equal to average',
    }[value]
  end

  # 2.5.161 Icv
  # The Icv structure specifies a color in the color table.
  # @param id [Integer]
  # @return [Hash]
  def _icv_d(id)
    result = Unxls::Biff8::Constants::ICV_COLOR_TABLE[id]

    result[:type] = case id
    when 0x0000..0x0007 then :builtin # The values that are greater than or equal to 0x0000 and less than or equal to 0x0007 specify built-in color constants. See table at p. 699
    when 0x0008..0x003F then :palette # The next 56 values in the table, the icv values greater than or equal to 0x0008 and less than or equal to 0x003F, specify the palette colors in the table. If a Palette record exists in this file, these icv values specify colors from the rgColor array in the Palette record. If no Palette record exists, these values specify colors in the default palette. See table at p. 699
    when 0x0040, 0x0041, 0x004D..0x004F, 0x0051, 0x7FFF then :display_settings # The remaining values in the color table specify colors associated with application display settings, see p. 701
    else raise "Unexpected icv value #{id}"
    end

    result
  end

  # 2.5.164 IcvXF
  # @param value [Integer]
  # @return [Integer]
  def icvxf(value)
    Unxls::BitOps.new(value).value_at(0..6)
  end

  # Unsigned 8 bit integer, little endian
  # @param data [String]
  # @return [Integer]
  def _int1b(data)
    data.unpack('C').first
  end

  # Unsigned 16 bit integer, little endian
  # @param data [String]
  # @return [Integer]
  def _int2b(data)
    data.unpack('v').first
  end

  # Unsigned 32 bit integer, little endian
  # @param data [String]
  # @return [Integer]
  def _int4b(data)
    data.unpack('V').first
  end

  # @param type [Symbol]
  # @param data [String]
  # @return [Integer, Symbol, String]
  def _interpret_xclr(type, data)
    case type
    when :XCLRINDEXED then icvxf(data.unpack('V').first) # See 2.5.164 IcvXF
    when :XCLRRGB then longrgba(data) # See 2.5.178 LongRGBA
    when :XCLRTHEMED then data.unpack('V').first # See 2.5.49 ColorTheme
    else data
    end
  end

  # 2.5.172 LEMMode
  # The LEMMode enumeration specifies the different edit modes for a table.
  # @param value [Integer]
  # @return [Symbol]
  def lemmode(value)
    {
      0x00 => :LEMNORMAL, # The table can be directly edited inline.
      0x01 => :LEMREFRESHCOPY, # The table is refreshed before editing is allowed because is it a copy of a table whose source is a Web based data provider list.
      0x02 => :LEMREFRESHCACHE, # The table is refreshed before editing is allowed because caching a user change failed.
      0x03 => :LEMREFRESHCACHEUNDO, # The table is refreshed before editing is allowed because undoing a cached user change failed.
      0x04 => :LEMREFRESHLOADED, # The table is refreshed before editing is allowed because on load the table source could not be re-connected.
      0x05 => :LEMREFRESHTEMPLATE, # The table is refreshed before editing is allowed because it was saved without having its data cached.
      0x06 => :LEMREFRESHREFRESH, # The table is refreshed before editing is allowed because a previous refresh failed.
      0x07 => :LEMNOINSROWSSPREQUIRED, # Rows cannot be inserted into this web based data provider list because there are hidden required columns.
      0x08 => :LEMNOINSROWSSPDOCLIB, # Rows cannot be inserted into this Web based data provider list because it is a document library.
      0x09 => :LEMREFRESHLOADDISCARDED, # The table is refreshed before editing is allowed because the user selected to discard cached changes upon loading.
      0x0A => :LEMREFRESHLOADHASHVALIDATION, # The table is refreshed before editing is allowed because the validation of the table's data area failed upon loading.
      0x0B => :LEMNOEDITSPMODVIEW, # Cannot allow the user to edit this table because of the type of moderated Web based data provider list it is.
    }[value]
  end

  # 2.5.174 List12BlockLevel
  # The List12BlockLevel structure specifies default block-level formatting information for a table, to be applied when the table expands.  Style gets applied before DXFN12List for each table region.
  # @param io [StringIO]
  # @return [String]
  def list12blocklevel(io)
    vars = io.read(36).unpack('l<*')

    result = {
      cbdxfHeader: vars[0], # cbdxfHeader (4 bytes): A signed integer that specifies the byte count for dxfHeader field. MUST be greater than or equal to zero.
      istnHeader: vars[1], # istnHeader (4 bytes): A signed integer that specifies a zero-based index to a Style record in the collection of Style records in the Globals Substream. The referenced Style specifies the cell style XF used for the table’s header row cells. If the value is -1, no style is specified for the table’s header row cells.
      cbdxfData: vars[2], # cbdxfData (4 bytes): A signed integer that specifies the byte count for dxfData field. MUST be greater than or equal to zero.
      istnData: vars[3], # istnData (4 bytes): A signed integer that specifies a zero-based index to a Style record in the collection of Style records in the Globals Substream. The referenced Style specifies the cell style used for the table’s data cells. If the value is -1, no style is specified for the table’s data cells.
      cbdxfAgg: vars[4], # cbdxfAgg (4 bytes): A signed integer that specifies the byte count for dxfAgg field. MUST be greater than or equal to zero.
      istnAgg: vars[5], # istnAgg (4 bytes): A signed integer that specifies a zero-based index to a Style record in the collection of Style records in the Globals Substream. The referenced Style specifies the cell style used for the table’s total row. If the value is -1, no style is specified for the table’s total row.
      cbdxfBorder: vars[6], # cbdxfBorder (4 bytes): A signed integer that specifies the byte count for dxfBorder field. MUST be greater than or equal to zero.
      cbdxfHeaderBorder: vars[7], # cbdxfHeaderBorder (4 bytes): A signed integer that specifies the byte count for dxfHeaderBorder field. MUST be greater than or equal to zero.
      cbdxfAggBorder: vars[8], # cbdxfAggBorder (4 bytes): A signed integer that specifies the byte count for dxfAggBorder field. MUST be greater than or equal to zero.
    }

    result[:dxfHeader] = dxfn12list(io.read(result[:cbdxfHeader])) unless result[:cbdxfHeader].zero? # dxfHeader (variable): An optional DXFN12List that specifies the formatting for the table’s header row cells. MUST exist if and only if cbdxfHeader is nonzero.
    result[:dxfData] = dxfn12list(io.read(result[:cbdxfData])) unless result[:cbdxfData].zero? # dxfData (variable): An optional DXFN12List that specifies the formatting for the table’s data cells. MUST exist if and only if cbdxfData is nonzero.
    result[:dxfAgg] = dxfn12list(io.read(result[:cbdxfAgg])) unless result[:cbdxfAgg].zero? # dxfAgg (variable): An optional DXFN12List that specifies the formatting for the table’s total row. MUST exist if and only if cbdxfAgg is nonzero.
    result[:dxfBorder] = dxfn12list(io.read(result[:cbdxfBorder])) unless result[:cbdxfBorder].zero? # dxfBorder (variable): An optional DXFN12 that specifies the formatting for the border of the table’s data cells. MUST exist if and only if cbdxfBorder is nonzero.
    result[:dxfHeaderBorder] = dxfn12list(io.read(result[:cbdxfHeaderBorder])) unless result[:cbdxfHeaderBorder].zero? # dxfHeaderBorder (variable): An optional DXFN12List that specifies the formatting for the border of the table’s header row cells. MUST exist if and only if cbdxfHeaderBorder is nonzero.
    result[:dxfAggBorder] = dxfn12list(io.read(result[:cbdxfAggBorder])) unless result[:cbdxfAggBorder].zero? # dxfAggBorder (variable): An optional DXFN12List that specifies the formatting for the border of the table’s total row. MUST exist if and only if cbdxfAggBorder is nonzero.

    result[:stHeader] = xlunicodestring(io) unless result[:istnHeader] == -1 # stHeader (variable): An optional XLUnicodeString that specifies the name of the style for the table’s header row cells. MUST exist if and only if istnHeader is not equal to -1. MUST be equal to the name of the Style record specified by istnHeader. If the style is a user-defined style, stHeader MUST be equal to the user field of the Style record.
    result[:stData] = xlunicodestring(io) unless result[:istnData] == -1 # stData (variable): An optional XLUnicodeString that specifies the name of the style for the table’s data cells. MUST exist if and only if istnData is not equal to -1. MUST be equal to the name of the Style record specified by istnData. If the style is a user-defined style, stData MUST be equal to the user field of the Style record.
    result[:stAgg] = xlunicodestring(io) unless result[:istnAgg] == -1 # stAgg (variable): An optional XLUnicodeString that specifies the name of the style for the table’s total row. MUST exist if and only if istnAgg is not equal to -1. MUST be equal to the name of the Style record specified by istnAgg. If the style is a user-defined style, stAgg MUST be equal to the user field of the Style record.

    result
  end

  # 2.5.175 List12DisplayName
  # The List12DisplayName structure specifies the name and comment strings for the table.
  # @param io [StringIO]
  # @return [String]
  def list12displayname(io)
    {
      stListName: xlunicodestring(io), # stListName (variable): An XLNameUnicodeString that specifies the table name. MUST be an empty string if the rgbName field of the TableFeatureType structure embedded in the Feature11 or Feature12 record that specifies the table is not empty. If the table name is not the same as the rgbName field of the TableFeatureType structure for this table, the table name is specified in stListName which is a case-insensitive unique name among all table names and defined names in the workbook.
      stListComment: xlunicodestring(io) # stListComment (variable): An XLUnicodeString that specifies a comment about the table.
    }
  end

  # 2.5.176 List12TableStyleClientInfo
  # The List12TableStyleClientInfo record specifies information about the style applied to a table.
  # @param io [StringIO]
  # @return [String]
  def list12tablestyleclientinfo(io)
    attrs = io.read(2).unpack('v').first
    attrs = Unxls::BitOps.new(attrs)

    {
      fFirstColumn: attrs.set_at?(0), # A - fFirstColumn (1 bit): A bit that specifies whether any table style elements (as specified by TableStyleElement) with a tseType field equal to 0x00000003 will be applied.
      fLastColumn: attrs.set_at?(1), # B - fLastColumn (1 bit): A bit that specifies whether any table style elements (as specified by TableStyleElement) with a tseType field equal to 0x00000004 will be applied.
      fRowStripes: attrs.set_at?(2), # C - fRowStripes (1 bit): A bit that specifies whether any table style elements (as specified by TableStyleElement) with a tseType field equal to 0x00000005 or 0x00000006 will be applied.
      fColumnStripes: attrs.set_at?(3), # D - fColumnStripes (1 bit): A bit that specifies whether any table style elements (as specified by TableStyleElement) with a tseType field equal to 0x00000007 or 0x00000008 will be applied.
      # E - unused1 (2 bits): Undefined and MUST be ignored.
      fDefaultStyle: attrs.set_at?(5), # F - fDefaultStyle (1 bit): A bit that specifies whether the style whose name is specified by stListStyleName is the default table style.
      # unused2 (9 bits): Undefined and MUST be ignored.
      stListStyleName: xlunicodestring(io) # stListStyleName (variable): An XLUnicodeString that specifies the name of the table style for the table. Length MUST be greater than zero and less than or equal to 255 characters. If the table style is a custom style, it is defined in a TableStyle record that has rgchName equal to this value.
    }
  end

  # @todo 2.5.198.19 ListParsedArrayFormula
  # The ListParsedArrayFormula structure specifies a formula (section 2.2.2) used in a table.
  # @param io [StringIO]
  # @raises [RuntimeError] .rgbextra
  def listparsedarrayformula(io)
    cce = io.read(2).unpack('v').first # cce (2 bytes): An unsigned integer that specifies the length of rgce in bytes. MUST be greater than 0.
    io.pos += cce # rgce (variable): An Rgce that specifies the sequence of Ptgs for the formula. MUST NOT contain PtgExp, PtgTbl, PtgElfLel, PtgElfRw, PtgElfCol, PtgElfRwV, PtgElfColV, PtgElfRadical, PtgElfRadicalS, PtgElfColS, PtgElfColSV, PtgElfRadicalLel, or PtgSxName.
    rgbextra(io) # rgcb (variable): An RgbExtra that specifies ancillary data for the formula.
  end

  # @todo 2.5.198.20 ListParsedFormula
  # The ListParsedFormula structure specifies a formula (section 2.2.2) used in a table.
  # @param io [StringIO]
  # @return [Symbol]
  def listparsedformula(io)
    cce = io.read(2).unpack('v').first # cce (2 bytes): An unsigned integer that specifies the length of rgce in bytes. MUST be greater than 0.
    io.pos += cce # rgce (variable): An Rgce that specifies the sequence of Ptgs for the formula. MUST NOT contain PtgExp, PtgTbl, PtgElfLel, PtgElfRw, PtgElfCol, PtgElfRwV, PtgElfColV, PtgElfRadical, PtgElfRadicalS, PtgElfColS, PtgElfColSV, PtgElfRadicalLel, or PtgSxName.

    :not_implemented
  end

  # 2.5.178 LongRGBA
  # The LongRGBA structure specifies a color as a combination of red, green, blue and alpha values.
  # @param data [String]
  # @return [Array<Symbol>]
  def longrgba(data)
    data.unpack('H*').first.upcase.to_sym
  end

  # 2.5.179 LPWideString
  # The LPWideString type specifies a Unicode string which is prefixed by a length.
  # @param data [String, StringIO]
  # @return [String]
  def lpwidestring(data)
    io = data.to_sio
    cch = io.read(2).unpack('v').first
    _read_unicodestring(io, cch, 1) # has double-byte characters
  end

  # See 2.5.113 Feat11FieldDataItem
  # @param value [Integer]
  # @return [Symbol]
  def _lfxidt_d(value)
    {
      0x1000 => :SOMITEM_SCHEMA,
      0x1001 => :SOMITEM_ATTRIBUTE,
      0x1002 => :SOMITEM_ATTRIBUTEGROUP,
      0x1003 => :SOMITEM_NOTATION,
      0x1100 => :SOMITEM_IDENTITYCONSTRAINT,
      0x1101 => :SOMITEM_KEY,
      0x1102 => :SOMITEM_KEYREF,
      0x1103 => :SOMITEM_UNIQUE,
      0x2000 => :SOMITEM_ANYTYPE,
      0x2100 => :SOMITEM_DATATYPE,
      0x2101 => :SOMITEM_DATATYPE_ANYTYPE,
      0x2102 => :SOMITEM_DATATYPE_ANYURI,
      0x2103 => :SOMITEM_DATATYPE_BASE64BINARY,
      0x2104 => :SOMITEM_DATATYPE_BOOLEAN,
      0x2105 => :SOMITEM_DATATYPE_BYTE,
      0x2106 => :SOMITEM_DATATYPE_DATE,
      0x2107 => :SOMITEM_DATATYPE_DATETIME,
      0x2108 => :SOMITEM_DATATYPE_DAY,
      0x2109 => :SOMITEM_DATATYPE_DECIMAL,
      0x210A => :SOMITEM_DATATYPE_DOUBLE,
      0x210B => :SOMITEM_DATATYPE_DURATION,
      0x210C => :SOMITEM_DATATYPE_ENTITIES,
      0x210D => :SOMITEM_DATATYPE_ENTITY,
      0x210E => :SOMITEM_DATATYPE_FLOAT,
      0x210F => :SOMITEM_DATATYPE_HEXBINARY,
      0x2110 => :SOMITEM_DATATYPE_ID,
      0x2111 => :SOMITEM_DATATYPE_IDREF,
      0x2112 => :SOMITEM_DATATYPE_IDREFS,
      0x2113 => :SOMITEM_DATATYPE_INT,
      0x2114 => :SOMITEM_DATATYPE_INTEGER,
      0x2115 => :SOMITEM_DATATYPE_LANGUAGE,
      0x2116 => :SOMITEM_DATATYPE_LONG,
      0x2117 => :SOMITEM_DATATYPE_MONTH,
      0x2118 => :SOMITEM_DATATYPE_MONTHDAY,
      0x2119 => :SOMITEM_DATATYPE_NAME,
      0x211A => :SOMITEM_DATATYPE_NCNAME,
      0x211B => :SOMITEM_DATATYPE_NEGATIVEINTEGER,
      0x211C => :SOMITEM_DATATYPE_NMTOKEN,
      0x211D => :SOMITEM_DATATYPE_NMTOKENS,
      0x211E => :SOMITEM_DATATYPE_NONNEGATIVEINTEGER,
      0x211F => :SOMITEM_DATATYPE_NONPOSITIVEINTEGER,
      0x2120 => :SOMITEM_DATATYPE_NORMALIZEDSTRING,
      0x2121 => :SOMITEM_DATATYPE_NOTATION,
      0x2122 => :SOMITEM_DATATYPE_POSITIVEINTEGER,
      0x2123 => :SOMITEM_DATATYPE_QNAME,
      0x2124 => :SOMITEM_DATATYPE_SHORT,
      0x2125 => :SOMITEM_DATATYPE_STRING,
      0x2126 => :SOMITEM_DATATYPE_TIME,
      0x2127 => :SOMITEM_DATATYPE_TOKEN,
      0x2128 => :SOMITEM_DATATYPE_UNSIGNEDBYTE,
      0x2129 => :SOMITEM_DATATYPE_UNSIGNEDINT,
      0x212A => :SOMITEM_DATATYPE_UNSIGNEDLONG,
      0x212B => :SOMITEM_DATATYPE_UNSIGNEDSHORT,
      0x212C => :SOMITEM_DATATYPE_YEAR,
      0x212D => :SOMITEM_DATATYPE_YEARMONTH,
      0x21FF => :SOMITEM_DATATYPE_ANYSIMPLETYPE,
      0x2200 => :SOMITEM_SIMPLETYPE,
      0x2400 => :SOMITEM_COMPLEXTYPE,
      0x4000 => :SOMITEM_PARTICLE,
      0x4001 => :SOMITEM_ANY,
      0x4002 => :SOMITEM_ANYATTRIBUTE,
      0x4003 => :SOMITEM_ELEMENT,
      0x4100 => :SOMITEM_GROUP,
      0x4101 => :SOMITEM_ALL,
      0x4102 => :SOMITEM_CHOICE,
      0x4103 => :SOMITEM_SEQUENCE,
      0x4104 => :SOMITEM_EMPTYPARTICLE,
      0x0800 => :SOMITEM_NULL,
      0x2800 => :SOMITEM_NULL_TYPE,
      0x4801 => :SOMITEM_NULL_ANY,
      0x4802 => :SOMITEM_NULL_ANYATTRIBUTE,
      0x4803 => :SOMITEM_NULL_ELEMENT,
    }[value]
  end

  # Map value to range -1.0..1.0
  # @param value [Integer] 16-bit signed int
  # @return [Float]
  def _map_tint(value)
    return 0.0 if value.zero? # save a few cycles
    result = (value < 0) ? (value / 32768.0) : (value / 32767.0)
    result.round(2)
  end

  # 2.5.186 NoteSh
  # The NoteSh structure specifies a comment associated with a cell.
  # @param io [StringIO]
  # @return [Hash]
  def notesh(io)
    row, col, attrs, id_obj = io.read(8).unpack('v4')
    attrs = Unxls::BitOps.new(attrs)
    {
      # changed to 'rw' for compability with cell fields
      rw: row, # row (2 bytes): A RW that specifies the row of the cell to which this comment is associated.
      col: col, # col (2 bytes): A Col that specifies the column of the cell to which this comment is associated.
      # A - reserved1 (1 bit): MUST be zero and MUST be ignored.
      fShow: attrs.set_at?(1), # B - fShow (1 bit): A bit that specifies whether the comment is shown at all times.
      # (C) reserved2, (D) unused1, (E) reserved3
      fRwHidden: attrs.set_at?(7), # F - fRwHidden (1 bit): A bit that specifies whether the row specified by row is hidden.
      fColHidden: attrs.set_at?(8), # G - fColHidden (1 bit): A bit that specifies whether the column specified by col is hidden.
      # reserved4 (7 bits): MUST be zero and MUST be ignored.
      idObj: id_obj, # idObj (2 bytes): An ObjId (2.5.188) that specifies the Obj record that specifies the comment text.
      stAuthor: xlunicodestring(io), # stAuthor (variable): An XLUnicodeString that specifies the name of the comment author. String length MUST be greater than or equal to 1 and less than or equal to 54.
      # unused2 (1 byte): Undefined and MUST be ignored.
    }
  end

  # 2.5.201 Phs
  # The Phs structure specifies the formatting information for a phonetic string.
  # @param data [String]
  # @return [Hash]
  def phs(data)
    ifnt, attrs = data.unpack('vv')
    attrs = Unxls::BitOps.new(attrs)

    ph_type = attrs.value_at(0..1)
    ph_type_d = {
      0x0 => :narrow_katakana,
      0x1 => :wide_katakana,
      0x2 => :hiragana,
      0x3 => :any # Use any type of characters as phonetic string
    }[ph_type]

    alc_h = attrs.value_at(2..3)
    alc_h_d = {
      0x0 => :general, # General alignment
      0x1 => :left, # Left aligned
      0x2 => :center, # Center aligned
      0x3 => :distributed # Distributed alignment
    }[alc_h]

    {
      ifnt: ifnt, # ifnt (2 bytes): A FontIndex structure that specifies the font.
      phType: ph_type, # A - phType (2 bits): An unsigned integer that specifies the type of the phonetic information.
      phType_d: ph_type_d,
      alcH: alc_h, # B - alcH (2 bits): An unsigned integer that specifies the alignment of the phonetic string.
      alcH_d: alc_h_d
      # unused (12 bits): Undefined and MUST be ignored.
    }
  end

  # 2.5.206 ReadingOrder
  # @param id [Integer]
  # @return [Symbol]
  def readingorder(id)
    {
      0 => :READING_ORDER_CONTEXT, # Context reading order
      1 => :READING_ORDER_LTR, # Left-to-right reading order
      2 => :READING_ORDER_RTL, # Right-to-left reading order
    }[id]
  end

  # @param record [Unxls::Biff8::Record]
  # @param cch [Integer]
  # @param high_byte [true, false]
  # @return [String]
  def _read_continued_string(record, cch, high_byte)
    cch_read = 0
    result = String.new

    # What Excel writes to cch is not really a number of characters in the string,
    # but rather a string bytesize / 2 if fHighByte is true, or just string bytesize if not,
    # so careful byte counting required when reading rgb field in case there are characters
    # encoded with more than 2 bytes (e.g. surrogate pairs):

    while cch > cch_read
      if record.bytes.eof?
        record.open_next_record_block
        attrs = record.bytes.read(1).unpack('C').first # fHighByte may be different in every Continue block
        high_byte = Unxls::BitOps.new(attrs).set_at?(0)
      end
      encoding, char_size = _encoding_params(high_byte)
      bytes_to_read = (cch - cch_read) * char_size
      string_bytes = record.bytes.read(bytes_to_read)
      cch_read += string_bytes.size / char_size
      result << string_bytes.force_encoding(encoding).encode(Encoding::UTF_8)
      break if record.end_of_data?
    end

    result
  end

  # @param io [StringIO]
  # @param cch [Integer]
  # @param flags [Integer]
  # @return [String]
  def _read_unicodestring(io, cch, flags)
    high_byte = Unxls::BitOps.new(flags).set_at?(0)
    encoding, char_size = _encoding_params(high_byte)
    _encode_string(io.read(cch * char_size), encoding)
  end

  # 2.5.208 Ref8
  # The Ref8 structure specifies a range of cells on the sheet.
  # @param data [String]
  # @return [Hash]
  def ref8(data)
    rw_first, rw_last, col_first, col_last = data.unpack('v4')

    # If rwFirst is 0 and rwLast is 0xFFFF, the specified range includes all the rows in the sheet
    # Same holds for columns
    {
      rwFirst: rw_first, # 0-based index of first row in the range
      rwLast: rw_last, # … last row in the range
      colFirst: col_first, # 0-based index of first column in the range
      colLast: col_last # … last column in the range
    }
  end

  # 2.5.209 Ref8U
  # The Ref8U structure specifies a range of cells on the sheet.
  # @param data [String]
  # @return [Hash]
  alias :ref8u :ref8 # structurally the same

  # @todo 2.5.198.103 RgbExtra
  # The RgbExtra structure specifies a set of structures, laid out sequentially in the file, that correspond to and MUST exist for certain Ptgs in the Rgce. The order of the structures MUST be the same as the order of the Ptgs in the Rgce that they correspond to.
  # @param io [StringIO]
  # @raises [RuntimeError]
  def rgbextra(io)
    # To find out the size of RgbExtra extra one must read all the ptgs, which requires
    # implementing formula parsing, which in turn is a fairly complex task (maybe in future).
    # But for now just
    raise 'Cannot parse RgbExtra structure yet, sorry.'
  end

  # 2.5.200 PhRuns
  # The PhRuns structure specifies a phonetic text run that is displayed above a text run.
  # @param io [StringIO]
  # @return [Hash]
  def rgphruns(io)
    ich_first, ich_mom, cch_mom = io.read(6).unpack('s<s<s<')

    {
      ichFirst: ich_first, # ichFirst (2 bytes): A signed integer that specifies the zero-based index of the first character of the phonetic text run in the rphssub.st field of the ExtRst structure that contains this PhRuns structure.
      ichMom: ich_mom, # ichMom (2 bytes): A signed integer that specifies the zero-based index of the first character of the text run
      cchMom: cch_mom # cchMom (2 bytes): A signed integer that specifies the count of characters in the text run specified in ichMom.
    }
  end

  # 2.5.217 RkNumber
  # @param value [Integer]
  # @return [Hash]
  def rknumber(value)
    data = Unxls::BitOps.new(value)

    f_x100 = data.set_at?(0) # num divided by 100?
    f_int = data.set_at?(1) # 0 - num is 64-bit floating point number, 1 - signed integer
    num_raw = ((value >> 2) << 2) # drop bits 0 and 1

    num = if f_int
      [num_raw].pack('V').unpack('l<').first # 32 bit signed little endian
    else
      num_to_64 = [num_raw].pack('xxxxV')
      num_to_64.unpack('E').first
    end

    f_x100 ? (num / 100.0) : num
  end

  # 2.5.218 RkRec
  # @param data [String]
  # @return [Hash]
  def rkrec(data)
    ixfe, rk_raw = data.unpack('vV')
    {
      ixfe: ixfe, # cell XF index
      RK: rknumber(rk_raw) # Numeric value
    }
  end

  # 2.5.219 RPHSSub
  # The RPHSSub structure specifies a phonetic string.
  # @param io [StringIO]
  # @return [Hash]
  def rphssub(io)
    crun, cch = io.read(4).unpack('vv')

    {
      crun: crun, # crun (2 bytes): An unsigned integer that specifies the number of phonetic text runs. MUST be less than or equal to 32767. If crun is zero, there is one phonetic text run.
      cch: cch, # cch (2 bytes): An unsigned integer that specifies the number of characters in the phonetic string.
      st: lpwidestring(io) # st (variable): An LPWideString that specifies the phonetic string. The character count in the string MUST be cch.
    }
  end

  # 2.5.226 Run
  # @param io [StringIO]
  # @return [Hash]
  def run(io)
    result = formatrun(io) # read 4 bytes
    io.read(4) # skip unused1, unused2
    result
  end

  # 2.5.232 Script
  # @param id [Integer]
  # @return [Symbol]
  def script(id)
    {
      0x00 => :SSSNONE, # Normal script
      0x01 => :SSSSUPER, # Superscript
      0x02 => :SSSSUB, # Subscript
      0xFF => :ignored, # Indicates that this specification is to be ignored
    }[id]
  end

  # 2.5.237 SharedFeatureType
  # The SharedFeatureType enumeration specifies the different types of Shared Features.
  # @param value [Integer]
  # @return [String]
  def sharedfeaturetype(value)
    {
      0x02 => :ISFPROTECTION, # Specifies the enhanced protection type. A Shared Feature of this type is used to protect a shared workbook by restricting access to the areas of the workbook and to the available functionality.
      0x03 => :ISFFEC2, # Specifies the ignored formula errors type. A Shared Feature of this type is used to specify the formula errors to be ignored.
      0x04 => :ISFFACTOID, # Specifies the smart tag type. A Shared Feature of this type is used to recognize certain types of entries (for example, proper names, dates/times, financial symbols) and flag them for action.
      0x05 => :ISFLIST, # Specifies the list type. A Shared Feature of this type is used to describe a table within a sheet.
    }[value]
  end

  # 2.5.240 ShortXLUnicodeString
  # The ShortXLUnicodeString structure specifies a Unicode string.
  # @param io [StringIO]
  # @return [String]
  def shortxlunicodestring(io)
    cch, flags = io.read(2).unpack('CC')
    _read_unicodestring(io, cch, flags)
  end

  # Signed 16 bit integer, native endian
  # @param data [String]
  # @return [Integer]
  def _sint2b(data)
    data.unpack('s').first
  end

  # 2.5.244 SourceType
  # The SourceType enumeration specifies the source type for a table.
  # @param value [Integer]
  # @return [Symbol]
  def sourcetype(value)
    {
      0x00 => :LTRANGE, # Range
      0x01 => :LTSHAREPOINT, # Read/write Web-based data provider list
      0x02 => :LTXML, # XML Mapper data
      0x03 => :LTEXTERNALDATA, # External data source (query table)<180>
    }[value]
  end

  # 2.5.246 SqRef
  # The SqRef structure specifies a sequence of Ref8 structures on the sheet.
  # @param io [StringIO]
  # @return [Hash]
  def sqref(io)
    cref = io.read(2).unpack('v').first

    {
      cref: cref, # cref (2 bytes): An unsigned integer that specifies the number of elements in rgrefs. MUST be less than or equal to 0x2000.
      rgrefs: cref.times.map { ref8u(io.read(8)) } # rgrefs (variable): An array of Ref8 structures. The number of elements in the array MUST be equal to cref.
    }
  end

  # 2.5.247 SqRefU
  alias :sqrefu :sqref

  # 2.5.248 Stxp, page 858
  # Specifies various formatting attributes of a font.
  # @param data [String]
  # @return [Hash]
  def stxp(data)
    twp_height, ts_raw, bls, sss, uls, b_family, b_char_set, _ = data.unpack('l<Vs<s<CCCC')

    {
      twpHeight: twp_height, # twpHeight (4 bytes): A signed integer that specifies the height of the font in twips. This value MUST be -1, 0, or between 20 and 8191. This value SHOULD NOT<181> be 0. A value of -1 specifies that this field is to be ignored.
      ts: ts(ts_raw), # ts (4 bytes): A Ts that specifies additional formatting attributes.
      bls: bls, # bls (2 bytes): A signed integer that specifies the font weight.
      bls_d: Unxls::Biff8::Structure.bold(bls),
      sss: sss, # sss (2 bytes): A signed integer that specifies whether the superscript or subscript or normal style of the font is used.
      sss_d: Unxls::Biff8::Structure.script(sss),
      uls: uls, # uls (1 byte): An unsigned integer that specifies the underline style.
      uls_d: Unxls::Biff8::Structure.underline(uls),
      bFamily: b_family, # bFamily (1 byte): An unsigned integer that specifies the font family, as defined by Windows API LOGFONT structure in [MSDN-FONTS].
      bFamily_d: Unxls::Biff8::Structure._font_family(b_family),
      bCharSet: b_char_set, # bCharSet (1 byte): An unsigned integer that specifies the character set, as defined by Windows API LOGFONT structure in [MSDN-FONTS].
      bCharSet_d: Unxls::Biff8::Structure._character_set(b_char_set),
      # unused (1 byte): Undefined and MUST be ignored.
    }
  end

  # 2.5.249 StyleXF
  # @param data [String]
  # @return [Hash]
  def stylexf(data)
    result = cellxf(data)
    # CellXF and StyleXF are identical except for the following fields which are not used in StyleXFs:
    %i(fAtrNum fAtrFnt fAtrAlc fAtrBdr fAtrPat fAtrProt fsxButton).each { |k| result.delete(k) }
    result
  end

  # 2.5.254 SXAxis
  # The SXAxis structure specifies the PivotTable axis referred to by the containing record.
  # @param value [Integer]
  # @return [Hash]
  def sxaxis4data(value)
    attrs = Unxls::BitOps.new(value)

    {
      sxaxisRw: attrs.set_at?(0), # A - sxaxisRw (1 bit): A bit that specifies whether this structure refers to the row axis.
      sxaxisCol: attrs.set_at?(1), # B - sxaxisCol (1 bit): A bit that specifies whether this structure refers to the column axis.
      sxaxisPage: attrs.set_at?(2), # C - sxaxisPage (1 bit): A bit that specifies whether this structure refers to the page axis.
      sxaxisData: attrs.set_at?(3), # D - sxaxisData (1 bit): A bit that specifies whether this structure refers to the value axis.
      # reserved (12 bits): MUST be zero, and MUST be ignored.
    }
  end

  # 2.5.266 TableFeatureType
  # The TableFeatureType structure specifies the definition of a table within a sheet.
  # @param io [StringIO]
  # @return [Hash]
  def tablefeaturetype(io)
    lt, id_list, crw_header, crw_totals, id_field_next, cb_fs_data, rup_build, _ = io.read(28).unpack('VVVVVVvv')

    result = {
      lt: lt, # lt (4 bytes): A SourceType that specifies the type of data source for the table.
      lt_d: sourcetype(lt),
      idList: id_list, # idList (4 bytes): An unsigned integer that specifies an identifier for the table. MUST be unique within the sheet. SHOULD<183> be unique within the workbook.
      crwHeader: crw_header == 1, # crwHeader (4 bytes): A Boolean (section 2.5.14) that specifies whether the table has a header row.
      crwTotals: crw_totals == 1, # crwTotals (4 bytes): A Boolean that specifies whether there is a total row.
      idFieldNext: id_field_next, # idFieldNext (4 bytes): An unsigned integer that specifies the next unique identifier to use when assigning unique identifiers to the fieldData.idField field of the table.
      cbFSData: cb_fs_data, # cbFSData (4 bytes): An unsigned integer that specifies the size, in bytes, of the fixed portion of this structure. The fixed portion starts at the lt field and ends at the rgbHashParam field. MUST be equal to 64.
      rupBuild: rup_build,  # rupBuild (2 bytes): An unsigned integer that specifies the build number of the application that wrote the structure.
      # unused1 (2 bytes): Undefined, and MUST be ignored.
    }

    attrs = Unxls::BitOps.new(io.read(4).unpack('V').first)
    result.merge!({
      # A - unused2 (1 bit): Undefined, and MUST be ignored.
      fAutoFilter: attrs.set_at?(1), # B - fAutoFilter (1 bit): A bit that specifies whether the table has an AutoFilter.  MUST be 1 when fPersistAutoFilter is 1.
      fPersistAutoFilter: attrs.set_at?(2), # C - fPersistAutoFilter (1 bit): A bit that specifies whether the AutoFilter is preserved for this table after data refresh operations.<184>
      fShowInsertRow: attrs.set_at?(3), # D - fShowInsertRow (1 bit): A bit that specifies whether the insert row is visible. MUST be 1 if fInsertRowInsCells is 1.
      fInsertRowInsCells: attrs.set_at?(4), # E - fInsertRowInsCells (1 bit): A bit that specifies whether rows below the table are shifted down because of the insert row being visible.
      fLoadPldwIdDeleted: attrs.set_at?(5), # F - fLoadPldwIdDeleted (1 bit): A bit that specifies whether the idDeleted field is present. MUST be zero if the lt field is not set to 0x00000001.
      fShownTotalRow: attrs.set_at?(6), # G - fShownTotalRow (1 bit): A bit that specifies whether the total row was ever visible.
      # H - reserved1 (1 bit):  MUST be zero and MUST be ignored.
      fNeedsCommit: attrs.set_at?(8), # I - fNeedsCommit (1 bit): A bit that specifies whether table modifications were not synchronized with the data source. MUST be zero if the lt field is not set to 0x00000001.
      fSingleCell: attrs.set_at?(9), # J - fSingleCell (1 bit): A bit that specifies whether the table is limited to a single cell. The table cannot have header rows, total rows, or multiple columns. If fSingleCell equals 1, the lt field MUST be set to 0x00000002.
      # K - reserved2 (1 bit):  MUST be zero and MUST be ignored.
      fApplyAutoFilter: attrs.set_at?(11), # L - fApplyAutoFilter (1 bit):  A bit that specifies whether the AutoFilter is currently applied. MUST be 1 if the AutoFilter is currently applied<185>.
      fForceInsertToBeVis: attrs.set_at?(12), # M - fForceInsertToBeVis (1 bit): A bit that specifies whether the insert row is forced to be visible because the table has no data.
      fCompressedXml: attrs.set_at?(13), # N - fCompressedXml (1 bit): A bit that specifies whether the cached data for this table in the List Data stream is compressed. MUST be zero if the lt field is not set to 0x00000001.
      fLoadCSPName: attrs.set_at?(14), # O - fLoadCSPName (1 bit): A bit that specifies whether the cSPName field is present. MUST be zero if the lt field is not set to 0x00000001.
      fLoadPldwIdChanged: attrs.set_at?(15), # P - fLoadPldwIdChanged (1 bit): A bit that specifies whether idChanged field is present. MUST be zero if the lt field is not set to 0x00000001.
      verXL: attrs.value_at(16..19), # verXL (4 bits): An unsigned integer that specifies the application version under which the table was created. MUST be either 0xB or 0xC<186>.
      fLoadEntryId: attrs.set_at?(20), # Q - fLoadEntryId (1 bit): A bit that specifies whether the entryId field is present.
      fLoadPllstclInvalid: attrs.set_at?(21), # R - fLoadPllstclInvalid (1 bit): A bit that specifies whether the cellInvalid field is present.  MUST be zero if the lt field is not set to 0x00000001.
      fGoodRupBld: attrs.set_at?(22), # S - fGoodRupBld (1 bit): A bit that specifies whether the rupBuild field is valid.
      # T - unused3 (1 bit): Undefined, and MUST be ignored.
      fPublished: attrs.set_at?(24), # U - fPublished (1 bit): A bit that specifies whether the table is published. This bit is ignored if the fPublishedBookItems field of the BookExt_Conditional12 structure is zero.
      # reserved3 (7 bits): Undefined, and MUST be ignored.
    })

    _, _, _, lem, _ = io.read(32).unpack('V3VC*')
    result.merge!({
      lPosStmCache: :not_implemented, # lPosStmCache (4 bytes): An unsigned integer that specifies the position of the cached data within the List Data stream.  Undefined and MUST be ignored if the lt field is not set to 0x00000001.
      cbStmCache: :not_implemented, # cbStmCache (4 bytes): An unsigned integer that specifies the size, in bytes, of the cached data within the List Data stream.  Undefined and MUST be ignored if the lt field is not set to 0x00000001.
      cchStmCache: :not_implemented, # cchStmCache (4 bytes): An unsigned integer that specifies the count of characters of the cached data within the List Data stream when the cached data is uncompressed.  Undefined and MUST be ignored if the lt field is not set to 0x00000001.
      LEMMode: lemmode(lem), # lem (4 bytes): A LEMMode enumeration that specifies the table edit mode. If lt is set to 0x00000000, 0x00000002 or 0x00000003, this field MUST be set to 0x00000000.
      rgbHashParam: :not_implemented, # rgbHashParam (16 bytes): An array of bytes that specifies round-trip information. SHOULD<187> be ignored and MUST be preserved if the lt field is set to 0x00000001. Undefined and MUST be ignored if the lt field is not set to 0x00000001.
    })

    result[:rgbName] = xlunicodestring(io) # rgbName (variable): An XLUnicodeString that specifies the name of the table. MUST be unique per workbook, and case-sensitive in all locales.
    result[:cFieldData] = io.read(2).unpack('v').first # cFieldData (2 bytes): An unsigned integer that specifies the number of columns in the table. MUST be greater than or equal to 0x0001 and less than or equal to 0x0100.

    if result[:fLoadCSPName]
      result[:cSPName] = xlunicodestring(io) # cSPName (variable): An XLUnicodeString that specifies the name of the cryptographic service provider used to specify rgbHashParam. This field is present only if fLoadCSPName is set to 1.
    end

    if result[:fLoadEntryId]
      result[:entryId] = xlunicodestring(io) # entryId (variable): An XLUnicodeString that specifies a unique identifier for the table. The string equals the value of the idList field, represented in decimal format, without any leading zeros. It is used when lt equals 0x00000002 and ignored otherwise. This field is present only if fLoadEntryId is set to 1.
    end

    result[:fieldData] = result[:cFieldData].times.map { feat11fielddataitem(io, result) } # fieldData (variable): An array of Feat11FieldDataItem that contains the specification of the columns of the table. The number of items in this array is specified by the cFieldData field.

    if result[:fLoadPldwIdDeleted]
      result[:idDeleted] = :not_implemented # idDeleted (variable): A Feat11RgSharepointIdDel structure that specifies the identifiers of deleted rows. This information is used when synchronizing with the Web based data provider’s data source. This field is only present if the fLoadPldwIdDeleted field is set to 1.
    end

    if result[:fLoadPldwIdChanged]
      result[:idChanged] = :not_implemented # idChanged (variable): A Feat11RgSharepointIdChange structure that specifies the identifiers of the edited rows. This information is used when synchronizing with the Web based data provider’s data source. This field is only present if the fLoadPldwIdChanged field is set to 1.
    end

    if result[:fLoadPllstclInvalid]
      result[:cellInvalid] = :not_implemented # cellInvalid (variable): A Feat11RgInvalidCells structure that specifies the location of cells within the table that contain values that are invalid based on validation rules on the Web based data provider. This field is only present if the fLoadPllstclInvalid field is set to 1.
    end

    result
  end

  # 2.5.270 Ts, page 879
  # The Ts structure specifies the italic and strikethrough formatting of a font.
  # @param value [Integer]
  # @return [Hash]
  def ts(value)
    a_unused3 = Unxls::BitOps.new(value)

    {
      # 0 (A) unused1 Undefined and MUST be ignored.
      ftsItalic: a_unused3.set_at?(1), # 1 (B) Specifies whether the text style is italic.
      # 2…6 unused2 (5 bits): Undefined and MUST be ignored.
      ftsStrikeout: a_unused3.set_at?(7) # 7 (C) Specifies whether the font has strikethrough formatting applied.
      # 8…31 unused3 (24 bits): Undefined and MUST be ignored.
    }
  end

  # See 2.4.321 TableStyleElement, p. 541
  # @param value [Integer]
  # @return [Hash]
  def _tse_type(value)
    {
      0x00 => :whole_table, # Whole table. If this table style is applied to a PivotTable view, this formatting type also applies to page field captions and page item captions.
      0x01 => :header_row, # Header row. If this table style is applied to a PivotTable view, this formatting type applies to the collection of rows above the data region. See S in the PivotTable Style Diagram.
      0x02 => :total_row, # Total row. If this table style is applied to a PivotTable view, this formatting type applies to the grand total row. See N in the PivotTable Style Diagram.
      0x03 => :first_column, # First column. If this table style is applied to a PivotTable view, this formatting type applies to the row label area, which can span multiple columns. See R in the PivotTable Style Diagram.
      0x04 => :last_column, # Last column. If this table style is applied to a PivotTable view, this formatting type applies to the grand total column. See A in the PivotTable Style Diagram.
      0x05 => :row_stripe_1, # Row stripe band 1
      0x06 => :row_stripe_2, # Row stripe band 2
      0x07 => :column_stripe_1, # Column stripe band 1
      0x08 => :column_stripe_2, # Column stripe band 2
      0x09 => :first_cell_header, # First cell of Header row. If this table style is applied to a PivotTable view, this formatting type applies to cells contained in area intersected by the header row and first column.
      0x0A => :last_cell_header, # Last cell of Header row. MUST be ignored if this table style is applied to a PivotTable view.
      0x0B => :first_cell_total, # First cell of Total row. MUST be ignored if this table style is applied to a PivotTable view.
      0x0C => :last_cell_total, # Last cell of Total row. MUST be ignored if this table style is applied to a PivotTable view.
      0x0D => :pt_outermost_subtotal_columns, # Outermost subtotal columns in a PivotTable view, specified by the columns displaying subtotals for the first Sxvd record in the PIVOTVD collection where the sxaxis field of the Sxvd record specifies the column axis. See B in the PivotTable Style Diagram. Used only for PivotTables.
      0x0E => :pt_alternating_even_subtotal_columns, # Alternating even subtotal columns in a PivotTable view, specified by the columns displaying subtotals for Sxvd records for which the zero-based index in the PIVOTVD collection is an odd number, omitting Sxvd records where the sxaxis field of the Sxvd record does not specify the column axis. See C in the PivotTable Style Diagram. Used only for PivotTables.
      0x0F => :pt_alternating_odd_subtotal_columns, # Alternating odd subtotal columns in a PivotTable view, specified by the columns displaying subtotals for Sxvd records for which the zero-based index in the PIVOTVD collection is an even number greater than zero, omitting Sxvd records where the sxaxis field of the Sxvd record does not specify the column axis. See D in the PivotTable Style Diagram. Used only for PivotTables.
      0x10 => :pt_outermost_subtotal_rows, # Outermost subtotal rows in a PivotTable view, specified by the rows displaying subtotals for the first Sxvd record in the PIVOTVD collection where the sxaxis field of the Sxvd record specifies the row axis. See M in the PivotTable Style Diagram. Used only for PivotTables.
      0x11 => :pt_alternating_even_subtotal_rows, # Alternating even subtotal rows in a PivotTable view, specified by the rows displaying subtotals for Sxvd records for which the zero-based index in the PIVOTVD collection is an odd number, omitting Sxvd records where the sxaxis field of the Sxvd record does not specify the row axis. See K in the PivotTable Style Diagram. Used only for PivotTables.
      0x12 => :pt_alternating_odd_subtotal_rows, # Alternating odd subtotal rows in a PivotTable view, specified by the rows displaying subtotals for Sxvd records for which the zero-based index in the PIVOTVD collection is an even number greater than zero, omitting Sxvd records where the sxaxis field of the Sxvd record does not specify the row axis. See J in the PivotTable Style Diagram. Used only for PivotTables.
      0x13 => :pt_empty_rows_after_each_subtotal_row, # Empty rows after each subtotal row. See L in the PivotTable Style Diagram. Used only for PivotTables.
      0x14 => :pt_outermost_column_subheadings, # Outermost column subheadings in a PivotTable view, specified by the columns displaying pivot field captions for the first Sxvd record in the PIVOTVD collection where the sxaxis field of the Sxvd record specifies the column axis. See O in the PivotTable Style Diagram. Used only for PivotTables.
      0x15 => :pt_alternating_even_column_subheadings, # Alternating even column subheadings in a PivotTable view, specified by the column columns displaying pivot field captions for Sxvd records for which the zero-based index in the PIVOTVD collection is an odd number, omitting Sxvd records where the sxaxis field of the Sxvd record does not specify the column axis. See P in the PivotTable Style Diagram. Used only for PivotTables.
      0x16 => :pt_alternating_odd_column_subheadings, # Alternating odd column subheadings in a PivotTable view, specified by the columns displaying pivot field captions for Sxvd records for which the zero-based index in the PIVOTVD collection is an even number greater than zero, omitting Sxvd records where the sxaxis field of the Sxvd record does not specify the column axis. See Q in the PivotTable Style Diagram. Used only for PivotTables.
      0x17 => :pt_outermost_row_subheadings, # Outermost row subheadings in a PivotTable view, specified by the rows displaying pivot field captions for the first Sxvd record in the PIVOTVD collection where the sxaxis field of the Sxvd record specifies the row axis. See G in the PivotTable Style Diagram. Used only for PivotTables.
      0x18 => :pt_alternating_even_row_subheadings, # Alternating even row subheadings in a PivotTable view, specified by the rows displaying pivot field captions for Sxvd records for which the zero-based index in the PIVOTVD collection is an odd number, omitting Sxvd records where the sxaxis field of the Sxvd record does not specify the row axis. See H in the PivotTable Style Diagram. Used only for PivotTables.
      0x19 => :pt_alternating_odd_row_subheadings, # Alternating odd row subheadings in a PivotTable view, specified by the rows displaying pivot field captions for Sxvd records for which the zero-based index in the PIVOTVD collection is an even number greater than zero, omitting Sxvd records where the sxaxis field of the Sxvd record does not specify the row axis. See I in the PivotTable Style Diagram. Used only for PivotTables.
      0x1A => :pt_page_field_captions, # Page field captions in a PivotTable view, specified by the cells displaying pivot field captions for the Sxvd records in the PIVOTVD collection where the sxaxis field of the Sxvd record specifies the page axis. See F in the PivotTable Style Diagram. Used only for PivotTables.
      0x1B => :pt_page_item_captions, # Page item captions in a PivotTable view, specified by the cells displaying pivot item captions for the Sxvd records in the PIVOTVD collection where the sxaxis field of the Sxvd record specifies the page axis. See E in the PivotTable Style Diagram. Used only for PivotTables.
    }[value]
  end

  # 2.5.274 Underline
  # @param id [Integer]
  # @return [Symbol]
  def underline(id)
    {
      0x00 => :ULSNONE, # No underline
      0x01 => :ULSSINGLE, # Single
      0x02 => :ULSDOUBLE, # Double
      0x21 => :ULSSINGLEACCOUNTANT, # Single accounting
      0x22 => :ULSDOUBLEACCOUNTANT, # Double accounting
      0xFF => :ignored # Indicates that this specification is to be ignored
    }[id]
  end

  # 2.5.275 VertAlign
  # @param id [Integer]
  # @return [Symbol]
  def vertalign(id)
    {
      0 => :ALCVTOP, # Top alignment
      1 => :ALCVCTR, # Center alignment
      2 => :ALCVBOT, # Bottom alignment
      3 => :ALCVJUST, # Justify alignment
      4 => :ALCVDIST, # Distributed alignment
    }[id]
  end

  # See p. 1087, notes 30-36
  # @param id [Integer]
  # @return [Symbol]
  def _verxlhigh(id)
    {
      0x0 => :'Excel 97',
      0x1 => :'Excel 2000',
      0x2 => :'Excel 2002',
      0x3 => :'Excel 2003',
      0x4 => :'Excel 2007',
      0x6 => :'Excel 2010',
      0x7 => :'Excel 2013'
    }[id]
  end

  alias :_verlastxlsaved :_verxlhigh

  # 2.5.279 XColorType
  # The XColorType enumeration specifies the color reference types.
  # @param id [Integer]
  # @return [Symbol]
  def xcolortype(id)
    {
      0 => :XCLRAUTO, # Automatic foreground/background colors (not specified)
      1 => :XCLRINDEXED, # xclrValue = index to palette color (IcvFX)
      2 => :XCLRRGB, # xclrValue = RGB color (LongRGBA)
      3 => :XCLRTHEMED, # xclrValue = index to Theme color (ColorTheme)
      4 => :XCLRNINCHED # Color not set
    }[id]
  end

  # Specifies the text alignment properties
  # 4-byte XF attribute structure common for:
  # 2.5.20 CellXF, page 597
  # 2.5.91 DXFALC, page 638
  # @param io [StringIO]
  # @return [Hash]
  def _xfalc(io)
    result = {}

    attrs = Unxls::BitOps.new(io.read(4).unpack('V').first)

    result[:alc] = (alc = attrs.value_at(0..2)) # alc (3 bits): A HorizAlign that specifies the horizontal alignment.
    result[:alc_d] = horizalign(alc)

    result[:fWrap] = attrs.set_at?(3) # A - fWrap (1 bit): A bit that specifies the text display when the text is wider than the cell.

    result[:alcV] = (alcv = attrs.value_at(4..6)) # alcV (3 bits): A VertAlign that specifies the vertical alignment.
    result[:alcV_d] = vertalign(alcv)

    result[:fJustLast] = attrs.set_at?(7) # B - fJustLast (1 bit): A bit that specifies whether the justified or distributed alignment of the cell is used on the last line of text (setting this to 1 is typical for East Asian text but not typical in other contexts).

    result[:trot] = (trot = attrs.value_at(8..15)) # trot (1 byte): An XFPropTextRotation that specifies the text rotation.
    result[:trot_d] = xfproptextrotation(trot)

    result[:cIndent] = attrs.value_at(16..19) # cIndent (4 bits): An unsigned integer that specifies the text indentation level.

    result[:fShrinkToFit] = attrs.set_at?(20) # C - fShrinkToFit (1 bit): A bit that specifies whether the cell is shrink to fit.

    # CellXF: D - reserved1 (1 bit): MUST be 0, and MUST be ignored.
    # DXFALC: D - fMergeCell (1 bit): A bit that specifies that the cell MUST be merged.
    result[:fMergeCell] = attrs.set_at?(21)

    result[:iReadingOrder] = (r_order = attrs.value_at(22..23)) # E - iReadingOrder (2 bits): A ReadingOrder that specifies the reading order.
    result[:iReadingOrder_d] = readingorder(r_order)

    # 24…25 reserved2 2 bits (F)

    # Do *not* update fields from parent XF?
    # ifmt
    result[:fAtrNum] = attrs.set_at?(26) # G - fAtrNum (1 bit): A bit that specifies that if the ifmt field of the XF record specified by the ixfParent field of the containing XF record is updated, the corresponding field of the containing XF record will *not* be set to the same value.
    # ifnt
    result[:fAtrFnt] = attrs.set_at?(27) # H - fAtrFnt (1 bit): A bit that specifies that if the ifnt field of the XF record specified by the ixfParent field of the containing XF record is updated, the corresponding field of the containing XF record will *not* be set to the same value.
    # alc, fWrap, alcV, fJustLast, trot, cIndent, fShrinkToFit, iReadingOrder
    result[:fAtrAlc] = attrs.set_at?(28) # I - fAtrAlc (1 bit): A bit that specifies that if the alc field, or the fWrap field, or the alcV field, or the fJustLast field, or the trot field, or the cIndent field, or the fShrinkToFit field or the iReadOrder field of the XF record specified by the ixfParent field of the containing XF record is updated, the corresponding fields of this structure will *not* be set to the same values.
    # dgLeft, dgRight, dgTop, dgBottom, dgDiag, icvLeft, icvRight, grbitDiag, icvTop, icvBottom, icvDiag
    result[:fAtrBdr] = attrs.set_at?(29) # J - fAtrBdr (1 bit): A bit that specifies that if the dgLeft field, or the dgRight field, or the dgTop field, or the dgBottom field, or the dgDiag field, or the icvLeft field, or the icvRight field, or the grbitDiag field, or the icvTop field, or the icvBottom field, or the icvDiag field of the XF record specified by the ixfParent field of the containing XF record is updated, the corresponding fields of this structure will *not* be set to the same values.
    # fls, icvFore, icvBack
    result[:fAtrPat] = attrs.set_at?(30) # K - fAtrPat (1 bit): A bit that specifies that if the fls field, the icvFore field, or the icvBack field of the XF record specified by the ixfParent field of the containing XF record is updated, the corresponding fields of this structure will *not* be set to the same values.
    # fLocked, fHidden
    result[:fAtrProt] = attrs.set_at?(31) # L - fAtrProt (1 bit): A bit that specifies that if the fLocked field or the fHidden field of the XF record specified by the ixfParent field of the containing XF record is updated, the corresponding fields of the containing XF record will *not* be set to the same values.

    result
  end

  # 8-byte XF attribute structure common for:
  # 2.5.20 CellXF, page 597
  # 2.5.92 DXFBdr, page 639
  # @param io [StringIO]
  # @return [Hash]
  def _xfbdr(io)
    result = {}

    attrs = io.read(4).unpack('V').first
    attrs = Unxls::BitOps.new(attrs)
    dg_left = attrs.value_at(0..3) # dgLeft (4 bits): A BorderStyle that specifies the logical left border formatting.
    result[:dgLeft] = dg_left
    result[:dgLeft_d] = borderstyle(dg_left)
    dg_right = attrs.value_at(4..7) # dgRight (4 bits): A BorderStyle that specifies the logical right border formatting.
    result[:dgRight] = dg_right
    result[:dgRight_d] = borderstyle(dg_right)
    dg_top = attrs.value_at(8..11) # dgTop (4 bits): A BorderStyle that specifies the top border formatting.
    result[:dgTop] = dg_top
    result[:dgTop_d] = borderstyle(dg_top)
    dg_bottom = attrs.value_at(12..15) # dgBottom (4 bits): A BorderStyle that specifies the bottom border formatting.
    result[:dgBottom] = dg_bottom
    result[:dgBottom_d] = borderstyle(dg_bottom)
    result[:icvLeft] = attrs.value_at(16..22) # icvLeft (7 bits): An unsigned integer that specifies the color of the logical left border. The value MUST be one of the values specified in IcvXF or 0. A value of 0 means the logical left border color has not been specified.
    result[:icvRight] = attrs.value_at(23..29) # icvRight (7 bits): An unsigned integer that specifies the color of the logical right border. The value MUST be one of the values specified in IcvXF or 0. A value of 0 means the logical right border color has not been specified.
    grbit_diag = attrs.value_at(30..31) # grbitDiag (2 bits): An unsigned integer that specifies which diagonal borders are present (if any).
    result[:grbitDiag] = grbit_diag
    result[:grbitDiag_d] = {
      0 => :'No diagonal border',
      1 => :'Diagonal-down border',
      2 => :'Diagonal-up border',
      3 => :'Both diagonal-down and diagonal-up',
    }[grbit_diag]

    attrs = io.read(4).unpack('V').first
    attrs = Unxls::BitOps.new(attrs)
    result[:icvTop] = attrs.value_at(0..6) # icvTop (7 bits): An unsigned integer that specifies the color of the top border. The value MUST be one of the values specified in IcvXF or 0. A value of 0 means the top border color has not been specified.
    result[:icvBottom] = attrs.value_at(7..13) # icvBottom (7 bits): An unsigned integer that specifies the color of the bottom border. The value MUST be one of the values specified in IcvXF or 0. A value of 0 means the bottom border color has not been specified.
    result[:icvDiag] = attrs.value_at(14..20) # icvDiag (7 bits): An unsigned integer that specifies the color of the diagonal border. The value MUST be one of the values specified in IcvXF or 0. A value of 0 means the diagonal border color has not been specified.
    dg_diag = attrs.value_at(21..24) # dgDiag (4 bits): A BorderStyle that specifies the diagonal border formatting.
    result[:dgDiag] = dg_diag
    result[:dgDiag_d] = borderstyle(dg_diag)
    result[:fHasXFExt] = attrs.set_at?(25) # fHasXFExt (1 bit): A bit that specifies whether an XFExt will extend the information in this XF.
    fls = attrs.value_at(26..31) # fls (6 bits): A FillPattern that specifies the fill pattern. If this value is 1, which specifies a solid fill pattern, then only icvFore is rendered.
    result[:fls] = fls
    result[:fls_d] = fillpattern(fls)

    result
  end

  # 2.5.280 XFExtGradient
  # The XFExtGradient structure specifies a gradient fill for a cell interior.
  # @param data [String]
  # @return [Hash]
  def xfextgradient(data)
    io = StringIO.new(data)

    result = xfpropgradient(io.read(44)) # gradient (44 bytes): An XFPropGradient that specifies the gradient fill.

    result[:cGradStops] = io.read(4).unpack('V').first # cGradStops (4 bytes): An unsigned integer that specifies the number of items in rgGradStops.

    result[:rgGradStops] = [] # rgGradStops (variable): An array of GradStop. Each array element specifies a gradient stop for this gradient fill.
    result[:cGradStops].times { result[:rgGradStops] << gradstop(io) }

    result
  end

  # See:
  # 2.4.355 XFExt, page 585
  # 2.5.281 XFExtNoFRT, page 885
  # @param io [StringIO]
  # @return [Hash]
  def _xfext(io)
    _, ixfe, _, cexts = io.read(8).unpack('v4')

    result = {
      # reserved1 (2 bytes): MUST be zero and MUST be ignored.
      ixfe: ixfe, # ixfe (2 bytes): An XFIndex structure that specifies the XF record in the file that this record extends.
      # reserved2 (2 bytes): MUST be zero and MUST be ignored.
      cexts: cexts, # cexts (2 bytes): An unsigned integer that specifies the number of elements in rgExt.
      rgExt: [] # rgExt (variable): An array of ExtProp. Each array element specifies a formatting property extension.
    }

    cexts.times { result[:rgExt] << extprop(io) }

    result
  end

  # 2.5.281 XFExtNoFRT
  # The XFExtNoFRT structure specifies a set of extensions to formatting properties.
  # @param io [StringIO]
  # @return [Hash]
  def xfextnofrt(io)
    result = _xfext(io)
    result.delete(:ixfe)
    result
  end

  # 2.5.283 XFProp
  # The XFProp structure specifies a formatting property.
  # @param io [StringIO]
  # @return [Hash]
  def xfprop(io)
    xf_prop_type, cb = io.read(4).unpack('vv')
    xf_prop_data_blob = io.read(cb - 4) # minus header length

    attrs = _xfproptype(xf_prop_type)
    method = attrs[:structure].downcase
    unpack_method = attrs[:unpack]
    data = unpack_method ? self.send(unpack_method, xf_prop_data_blob) : xf_prop_data_blob

    result = {
      xfPropType: xf_prop_type, # xfPropType (2 bytes): An unsigned integer that specifies the type of the formatting property.
      cb: cb, # cb (2 bytes): An unsigned integer that specifies the size of this XFProp structure.
      _description: attrs[:description],
      _structure: attrs[:structure],
      xfPropDataBlob_d: self.send(method, data) # xfPropDataBlob (variable): A field that specifies the formatting property data.
    }

    result[:xfPropDataBlob] = xf_prop_data_blob if unpack_method

    result
  end

  # 2.5.284 XFPropBorder
  # @param data [String]
  # @return [Hash]
  def xfpropborder(data)
    io = StringIO.new(data)
    color_data = io.read(8)
    dg_border = io.read(2).unpack('v').first

    {
      color: xfpropcolor(color_data),
      dgBorder: dg_border,
      dgBorder_d: borderstyle(dg_border)
    }
  end

  # 2.5.285 XFPropColor
  # The XFPropColor structure specifies a color.
  # @param data [String]
  # @return [Hash]
  def xfpropcolor(data)
    io = StringIO.new(data)
    attrs, icv, n_tint_shade = io.read(4).unpack('CCs<')
    attrs = Unxls::BitOps.new(attrs)
    xclr_type = attrs.value_at(1..7)

    {
      fValidRGBA: attrs.set_at?(0), # A - fValidRGBA (1 bit): A bit that specifies whether the xclrType, icv and nTintShade fields were used to set the dwRgba field.
      xclrType: xclr_type, # xclrType (7 bits): An XColorType that specifies how the color information is stored.
      xclrType_d: xcolortype(xclr_type),
      icv: icv, # icv (1 byte): An unsigned integer that specifies color information. If xclrType equals 0x01, this field MUST be one of the values specified in IcvXF, or equal 0. If xclrType equals 0x03, this field MUST be one of the values specified in ColorTheme.
      nTintShade: n_tint_shade, # nTintShade (2 bytes): A signed integer that specifies the tint of the color. This value is mapped to the range -1.0 to 1.0. Positive values lighten the color, and negative values darken the color.
      nTintShade_d: _map_tint(n_tint_shade),
      dwRgba: longrgba(io.read(4)) # dwRgba (4 bytes): A LongRGBA that specifies the color.
    }
  end

  # 2.5.286 XFPropGradient
  # The XFPropGradient structure specifies a gradient fill.
  # @param data [String]
  # @return [Hash]
  def xfpropgradient(data)
    io = StringIO.new(data)

    type = io.read(4).unpack('V').first rescue binding.pry
    type_d = {
      0 => :linear,
      1 => :rectangular
    }[type]

    {
      type: type, # type (4 bytes): A Boolean (section 2.5.14) that specifies the gradient type.
      type_d: type_d,
      numDegree: xnum(io.read(8)).round(2), # numDegree (8 bytes): An Xnum (section 2.5.342) that specifies the gradient angle in degrees for a linear gradient. The gradient angle specifies the angle at which gradient strokes are drawn.
      numFillToLeft: xnum(io.read(8)).round(2), # numFillToLeft (8 bytes): An Xnum that specifies the left coordinate of the inner rectangle for a rectangular gradient, where (0.0,0.0) is the upper-left hand corner of the inner rectangle.
      numFillToRight: xnum(io.read(8)).round(2), # numFillToRight (8 bytes): An Xnum that specifies the right coordinate of the inner rectangle for a rectangular gradient, where (0.0,0.0) is the upper-left hand corner of the inner rectangle.
      numFillToTop: xnum(io.read(8)).round(2), # numFillToTop (8 bytes): An Xnum that specifies the top coordinate of the inner rectangle for a rectangular gradient, where (0.0,0.0) is the upper-left hand corner of the inner rectangle.
      numFillToBottom: xnum(io.read(8)).round(2), # numFillToBottom (8 bytes): An Xnum that specifies the bottom coordinate of the inner rectangle for a rectangular gradient, where (0.0,0.0) is the upper-left hand corner of the inner rectangle.
    }
  end

  # 2.5.287 XFPropGradientStop
  # @param data [String]
  # @return [Hash]
  def xfpropgradientstop(data)
    io = StringIO.new(data)
    io.read(2) # unused

    {
      numPosition: xnum(io.read(8)),
      color: xfpropcolor(io.read(8))
    }
  end

  # 2.5.288 XFProps
  # This structure specifies an array of formatting properties.
  # @param io [StringIO]
  # @return [Hash]
  def xfprops(io)
    _, cprops = io.read(4).unpack('vv')

    result = {
      # reserved (2 bytes): MUST be zero and MUST be ignored.
      cprops: cprops, # cprops (2 bytes): An unsigned integer that specifies the number of XFProp structures in xfPropArray.
      xfPropArray: [] # xfPropArray (variable): An array of XFProp. Each array element specifies a formatting property. The array of properties specifies the full set of formatting properties.
    }

    cprops.times { result[:xfPropArray] << xfprop(io) }

    result
  end

  # 2.5.289 XFPropTextRotation
  # @param value [Integer]
  # @return [Hash]
  def xfproptextrotation(value)
    case value
    when 0..90 then :counterclockwise
    when 91..180 then :clockwise
    when 255 then :vertical
    else nil
    end
  end

  # See 2.5.283 XFProp page 887
  # @param id [Integer]
  # @return [Hash]
  def _xfproptype(id)
    {
      0x0000 => { description: :fill_pattern,         structure: :FillPattern, unpack: :_int1b },
      0x0001 => { description: :foreground_color,     structure: :XFPropColor },
      0x0002 => { description: :background_color,     structure: :XFPropColor },
      0x0003 => { description: :gradient_fill,        structure: :XFPropGradient },
      0x0004 => { description: :gradient_stop,        structure: :XFPropGradientStop },
      0x0005 => { description: :text_color,           structure: :XFPropColor },
      0x0006 => { description: :top_border,           structure: :XFPropBorder },
      0x0007 => { description: :bottom_border,        structure: :XFPropBorder },
      0x0008 => { description: :left_border,          structure: :XFPropBorder },
      0x0009 => { description: :right_border,         structure: :XFPropBorder },
      0x000A => { description: :diagonal_border,      structure: :XFPropBorder },
      0x000B => { description: :vertical_border,      structure: :XFPropBorder },
      0x000C => { description: :horizontal_border,    structure: :XFPropBorder },
      0x000D => { description: :diagonal_up,          structure: :_bool1b }, # 1 == used, 0 == not used
      0x000E => { description: :diagonal_down,        structure: :_bool1b },
      0x000F => { description: :horizontal_alignment, structure: :HorizAlign, unpack: :_int1b },
      0x0010 => { description: :vertical_alignment,   structure: :VertAlign, unpack: :_int1b },
      0x0011 => { description: :text_rotation,        structure: :XFPropTextRotation, unpack: :_int1b },
      0x0012 => { description: :indentation_level,    structure: :_int2b }, # 2-byte integer
      0x0013 => { description: :reading_order,        structure: :ReadingOrder, unpack: :_int1b },
      0x0014 => { description: :text_wrap,            structure: :_bool1b },
      0x0015 => { description: :justify,              structure: :_bool1b },
      0x0016 => { description: :shrink_to_fit,        structure: :_bool1b },
      0x0017 => { description: :cell_merged,          structure: :_bool1b },
      0x0018 => { description: :font_name,            structure: :LPWideString },
      0x0019 => { description: :font_weight,          structure: :Bold, unpack: :_int2b },
      0x001A => { description: :underline_style,      structure: :Underline, unpack: :_int2b },
      0x001B => { description: :script_style,         structure: :Script, unpack: :_int2b },
      0x001C => { description: :italic,               structure: :_bool1b },
      0x001D => { description: :strikethrough,        structure: :_bool1b },
      0x001E => { description: :outline,              structure: :_bool1b },
      0x001F => { description: :shadow,               structure: :_bool1b },
      0x0020 => { description: :condensed,            structure: :_bool1b },
      0x0021 => { description: :extended,             structure: :_bool1b },
      0x0022 => { description: :character_set,        structure: :_character_set, unpack: :_int1b },
      0x0023 => { description: :font_family,          structure: :_font_family, unpack: :_int1b },
      0x0024 => { description: :text_size,            structure: :_int4b },
      0x0025 => { description: :font_scheme,          structure: :FontScheme },
      0x0026 => { description: :number_format,        structure: :XLUnicodeString },
      0x0029 => { description: :ifmt,                 structure: :_int2b },
      0x002A => { description: :relative_indentation, structure: :_sint2b },
      0x002B => { description: :locked,               structure: :_bool1b },
      0x002C => { description: :hidden,               structure: :_bool1b },
    }[id]
  end

  # 2.5.294 XLUnicodeString
  # The XLUnicodeString structure specifies a Unicode string.
  # @param data [String, StringIO]
  # @return [String]
  def xlunicodestring(data)
    io = data.to_sio
    cch, flags = io.read(3).unpack('vC')
    _read_unicodestring(io, cch, flags)
  end

  # 2.5.296 XLUnicodeStringNoCch
  # The XLUnicodeStringNoCch structure specifies a Unicode string.
  # @param io [StringIO]
  # @param cch [Integer]
  # @return [String]
  def xlunicodestringnocch(io, cch)
    flags = io.read(1).unpack('C').first
    _read_unicodestring(io, cch, flags)
  end

  # 2.5.293 XLUnicodeRichExtendedString
  # The XLUnicodeRichExtendedString structure specifies a Unicode string, which can contain formatting information and phonetic string data.
  # @param record [Unxls::Biff8::Record] Has SST and its Continue records
  # @return [Hash]
  def xlunicoderichextendedstring(record)
    record.open_next_record_block if record.bytes.eof?

    cch, attrs = record.bytes.read(3).unpack('vC')
    attrs = Unxls::BitOps.new(attrs)
    result = {
      cch: cch, # cch (2 bytes): An unsigned integer that specifies the count of characters in the string.
      fHighByte: attrs.set_at?(0), # A - fHighByte (1 bit): A bit that specifies whether the characters in rgb are double-byte characters.
      # B - reserved1 (1 bit): MUST be zero, and MUST be ignored.
      fExtSt: attrs.set_at?(2), # C - fExtSt (1 bit): A bit that specifies whether the string contains phonetic string data.
      fRichSt: attrs.set_at?(3), # D - fRichSt (1 bit): A bit that specifies whether the string is a rich string and the string has at least two character formats applied.
      # reserved2 (4 bits): MUST be zero, and MUST be ignored.
    }

    if result[:fRichSt]
      result[:cRun] = record.bytes.read(2).unpack('v').first # cRun (2 bytes): An optional unsigned integer that specifies the number of elements in rgRun.
    end

    if result[:fExtSt]
      result[:cbExtRst] = record.bytes.read(4).unpack('l<').first # cbExtRst (4 bytes): An optional signed integer that specifies the byte count of ExtRst.
    end

    result[:rgb] = _read_continued_string(record, cch, result[:fHighByte]) # rgb (variable): An array of bytes that specifies the characters in the string.

    if result[:fRichSt]
      result[:rgRun] = []
      result[:cRun].times do
        record.open_next_record_block if record.bytes.eof?
        result[:rgRun] << formatrun(record.bytes) # rgRun (variable): An optional array of FormatRun structures that specifies the formatting for each text run.
        break if record.end_of_data?
      end
    end

    if result[:fExtSt]
      result[:ExtRst] = String.new
      while (unread = result[:cbExtRst] - result[:ExtRst].size) > 0
        record.open_next_record_block if record.bytes.eof?
        result[:ExtRst] << record.bytes.read(unread)
        break if record.end_of_data?
      end
      result[:ExtRst] = extrst(result[:ExtRst]) # ExtRst (variable): An optional ExtRst that specifies the phonetic string data. The size of this field is cbExtRst.
    end

    result
  end

  # 2.5.342 Xnum
  # Xnum is a 64-bit binary floating-point number as specified in [IEEE754]. This value MUST NOT be infinity, denormalized, not-a-number (NaN), nor negative zero.
  # @param data [String]
  # @return [Float]
  def xnum(data)
    data.unpack('E').first
  end

  # 2.5.343 XORObfuscation
  # The XORObfuscation structure specifies the XOR obfuscation.
  # @param io [StringIO]
  # @return [Hash]
  def xorobfuscation(io)
    key, verification_bytes = io.read(4).unpack('vv')

    {
      key: key, # key (2 bytes): An unsigned integer that specifies the obfuscation key. See [MS-OFFCRYPTO], 2.3.6.2 section, the first step of initializing XOR array where it describes the generation of 16-bit XorKey value.
      verificationBytes: verification_bytes # verificationBytes (2 bytes): An unsigned integer that specifies the password verification identifier. See Password Verifier Algorithm.
    }
  end

end
