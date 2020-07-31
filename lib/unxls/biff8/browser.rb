# frozen_string_literal: true

class Unxls::Biff8::Browser

  attr_reader :parsed

  # @param file [String, Pathname, File]
  # @param settings [Hash]
  def initialize(file, settings = {})
    @file = Pathname.new(file).to_s
    @settings = settings
    reload
  end

  # @param return_self [true, false] (true)
  # @return [Unxls::Biff8::Browser, nil]
  def reload(return_self = true)
    @parsed = Unxls.parse(@file, @settings)
    self if return_self
  end

  # WorkBookStream
  # @return [Hash]
  def wbs
    @parsed[:workbook_stream]
  end

  # Workbook stream, globals substream
  # @return [Hash]
  def globals
    raise unless wbs.first[:BOF][:dt_d] == :globals
    wbs.first
  end

  # Cell address to location index
  # @return [Hash]
  def cell_index
    wbs.last
  end

  # Map certain records, or their non-nil attributes, or anything non-nil that the passed block will return.
  #
  # @example
  #   # Collect all records of certain type:
  #   Browser.new(parsed).mapif(:Font)
  #   #=> [{ dyHeight: … }, …]
  #
  #   # Collect certain attributes:
  #   Browser.new(parsed).mapif(:Font) { |record| record[:fontName] }
  #   #=> ['Arial', …]
  #
  #   # Index is passed into the block:
  #   Browser.new(parsed).mapif(:Font) do |r, index|
  #     { index => r[:fontName] }
  #   end
  #   #=> [{ 0 => 'Arial' }, … ]
  #
  #   # Wrap in other value to collect nils:
  #   Browser.new(parsed).mapif(:SomeRecord) { |r| [r[:someField]] }
  #   #=> [[nil], [18], [23], …]
  #
  #   # Collect only those records that have attributes of certain value:
  #   Browser.new(parsed).mapif(:Blank, :LabelSst) { |r| r if r[:ixfe] == 42 }
  #   #=> [{ ifxe: 42, _record: { name: :Blank, … } }, …]
  #
  #   # Join other records. Browser instance and substream index are passed into the block as well for unlimited power:
  #   Browser.new(parsed).mapif(:XF) do |r1, i1, browser, substream_index|
  #     { [substream_index, i1] => browser.mapif(:Style) { |r2| r2 if r2[:ixfe] == i1 } }
  #   end
  #   #=> [
  #   #=>   { [0, 0] => { ixfe: 0, fBuiltIn: true, … } },
  #   #=>   { [0, 1] => [] },
  #   #=>   …
  #   #=> ]
  #
  # @param *record_types [Array<Symbol>] list of types of records to map
  # @param first: [TrueClass, FalseClass] false Only return the first match
  # @return [Array, Object, nil]
  def mapif(*record_types, first: false)
    result = []

    wbs.each_with_index do |substream, substream_index|
      record_types.each do |rt|
        [substream[rt]].flatten.compact.each_with_index do |record, index|
          if block_given?
            temp = yield(record, index, self, substream_index)
            next if temp.nil?
            return(temp) if first
            result << temp
          else
            result << record
          end
        end
      end
    end

    first ? result.first : result
  end

  # @param sheet_name [String, Symbol]
  # @return [Integer, nil]
  def get_substream_index(sheet_name)
    ref = globals[:BoundSheet8].find { |s| s[:stName] == sheet_name.to_s }
    return unless ref

    wbs.find_index do |s|
      s[:BOF][:_record][:pos] == ref[:lbPlyPos] if s[:BOF] && s[:BOF][:dt_d] != :globals
    end
  end

  # @param name [String, Symbol]
  # @return [Hash, nil]
  def get_sheet(name)
    sheet_index = get_substream_index(name)
    wbs[sheet_index] if sheet_index
  end

  # @param substream_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @param :record_type [Symbol] (nil)
  # @return [Hash, nil]
  def get_cell(substream_index, row, column, record_type: nil)
    address_key = Unxls::Biff8::WorkbookStream.make_address_key(substream_index, row, column)
    return unless (access_address = cell_index[address_key])
    inferred_record_type, record_index = access_address.to_s.split('_')
    wbs[substream_index][record_type || inferred_record_type.to_sym][record_index.to_i]
  end

  # @param font_index [Integer]
  # @return [Hash, nil]
  def get_font(font_index)
    font_index -= 1 if font_index >= 4 # See 2.5.129 FontIndex, p. 677
    globals[:Font][font_index]
  end

  # @param substream_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @return [Hash, nil]
  def get_format(substream_index, row, column)
    return unless (xf = get_xf(substream_index, row, column))

    format_index = xf[:ifmt]
    globals[:Format].find { |r| r[:ifmt] == format_index }
  end

  # @param substream_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @return [Hash, nil]
  def get_sst(substream_index, row, column)
    sst_index = get_cell(substream_index, row, column)[:isst]
    globals[:SST][:rgb][sst_index] if sst_index
  end

  # @param substream_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @return [Hash, nil]
  def get_xf(substream_index, row, column)
    return unless (cell = get_cell(substream_index, row, column))
    globals[:XF][cell_xf_index(cell, column)]
  end

  # @param substream_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @return [Hash, nil]
  def get_xfext(substream_index, row, column)
    return unless (cell = get_cell(substream_index, row, column))
    ixfe = cell_xf_index(cell, column)
    globals[:XFExt].find { |r| r[:ixfe] == ixfe } # @todo speedup with xf_xfext index?
  end

  def get_style(substream_index, row, column)
    return unless (xf = get_xf(substream_index, row, column))
    cell_parent_xf_ixfe = xf[:ixfParent]
    globals[:Style].find { |r| r[:ixfe] == cell_parent_xf_ixfe }
  end

  def get_styleext(substream_index, row, column)
    style = get_style(substream_index, row, column)
    globals[:StyleExt][style[:_record][:index]]
  end

  def get_tablestyle(tablestyle_name)
    globals[:TableStyle].find { |r| r[:rgchName] == tablestyle_name }
  end

  def get_tablestyleelement(tablestyle_name, type)
    index = get_tablestyle(tablestyle_name)[:_record][:index]
    globals[:TableStyleElement].find { |r| r[:_tsi] == index && r[:tseType_d] == type }
  end

  # DXF related to TableStyleElement record
  # @param tablestyle_name [String]
  # @param type [Symbol]
  # @return [Hash]
  def get_tse_dxf(tablestyle_name, type)
    globals[:DXF][ get_tablestyleelement(tablestyle_name, type)[:index] ]
  end

  # @param substream_index [Integer]
  # @param column [Integer]
  # @return [Hash, nil]
  def get_colinfo(substream_index, column)
    wbs[substream_index][:ColInfo].find { |r| column >= r[:colFirst] && column <= r[:colLast] }
  end

  # @param substream_index [Integer]
  # @param row [Integer]
  # @return [Hash, nil]
  def get_row(substream_index, row)
    wbs[substream_index][:Row].find { |r| r[:rw] == row }
  end

  # @param substream_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @return [Hash, nil]
  def get_note(substream_index, row, column)
    address_key = Unxls::Biff8::WorkbookStream.make_address_key(substream_index, row, column)
    return unless (record_index = cell_index[:notes][address_key])
    wbs[substream_index][:Note][record_index]
  end
  
  # @param substream_index [Integer]
  # @param object_id [Integer]
  # @return [Hash, nil]
  def get_obj(substream_index, object_id)
    wbs[substream_index][:Obj].find { |r| r[:cmo][:id] == object_id }
  end

  # @param substream_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @return [Hash, nil]
  def get_hlink(substream_index, row, column)
    address_key = Unxls::Biff8::WorkbookStream.make_address_key(substream_index, row, column)
    return unless (record_index = cell_index[:hlinks][address_key])
    wbs[substream_index][:HLink][record_index]
  end

  # @param substream_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @return [Hash, nil]
  def get_hlinktooltip(substream_index, row, column)
    address_key = Unxls::Biff8::WorkbookStream.make_address_key(substream_index, row, column)
    return unless (record_index = cell_index[:hlinktooltips][address_key])
    wbs[substream_index][:HLinkTooltip][record_index]
  end

  # @todo @debug
  # @param substream_index [Integer]
  # @param row [Integer]
  # @param column [Integer]
  # @return [Hash, nil]
  def get_cell_styling(substream_index, row, column) # @todo -> xf_record_chain ?
    return unless (cell = get_cell(substream_index, row, column))
    cell_ixfe = cell_xf_index(cell, column)

    cell_xf = globals[:XF][cell_ixfe]
    cell_xfext = globals[:XFExt].find { |r| r[:ixfe] == cell_ixfe } # @todo index?

    cell_parent_xf_ixfe = cell_xf[:ixfParent]

    style_xf = globals[:XF][cell_parent_xf_ixfe]
    style_xfext = globals[:XFExt].find { |r| r[:ixfe] == cell_parent_xf_ixfe } # @todo index?

    style = globals[:Style].find { |r| r[:ixfe] == cell_parent_xf_ixfe } # @todo index?
    styleext = globals[:StyleExt][style[:_record][:index]]

    result = {
      xf: cell_xf,
      xfext: cell_xfext,
      style_xf: style_xf,
      style_xfext: style_xfext,
      style: style,
      styleext: styleext,
    }

    cell_xf_font_index = cell_xf[:ifnt]
    cell_xf_font_index -= 1 if cell_xf_font_index >= 4 # See 2.5.129 FontIndex, p. 677
    result[:xf_font] = globals[:Font][cell_xf_font_index]

    if (style_xf_font_index = style_xf[:ifnt]) != cell_xf[:ifnt]
      style_xf_font_index -= 1 if style_xf_font_index >= 4 # See 2.5.129 FontIndex, p. 677
      result[:style_xf_font] = globals[:Font][style_xf_font_index]
    end

    cell_xf_format_index = cell_xf[:ifmt]
    result[:xf_format] = globals[:Format].find { |r| r[:ifmt] == cell_xf_format_index }

    if (style_xf_format_index = style_xf[:ifmt]) != cell_xf_format_index
      result[:style_xf_format] = globals[:Format].find { |r| r[:ifmt] == style_xf_format_index }
    end

    if (rows = wbs[substream_index][:Row])
      row_record = rows.find { |r| r[:rw] == row }
      result[:row] = row_record if row_record
    end

    if (columns = wbs[substream_index][:ColInfo])
      column_record = columns.find { |r| (r[:colFirst]..r[:colLast]).include?(column) }
      result[:column] = column_record if column_record
    end

    result
  end

  # @todo @debug
  def xf_styles
    mapif(:XF) do |r, i, s|
      style = s.mapif(:Style) { |r1| r1 if r1[:ixfe] == i  }.first
      {
        index:     i,
        fStyle:    r[:fStyle],
        _builtin:  r[:_description],
        ixfParent: r[:ixfParent],
        Style:     style
      } if style
    end
  end

  private

  # @param cell [Hash]
  # @param column [Integer]
  # @return [Integer]
  def cell_xf_index(cell, column)
    case (record_type = cell[:_record][:name])
    when :LabelSst, :RK, :Blank, :Formula, :Number, :BoolErr, :Label
      cell[:ixfe]
    when :MulRk
      rgkrec_index = column - cell[:colFirst]
      cell[:rgrkrec][rgkrec_index][:ixfe]
    when :MulBlank
      rgixfe_index = column - cell[:colFirst]
      cell[:rgixfe][rgixfe_index]
    else
      raise "Dunno how to process record type :#{record_type}"
    end
  end

end