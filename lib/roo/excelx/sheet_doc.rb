require 'roo/excelx/extractor'

module Roo
  class Excelx
    class SheetDoc < Excelx::Extractor
      def initialize(path, relationships, styles, shared_strings, workbook, options = {})
        super(path)
        @options = options
        @relationships = relationships
        @styles = styles
        @shared_strings = shared_strings
        @workbook = workbook
      end

      def cells(relationships)
        @cells ||= extract_cells(relationships)
      end

      def hyperlinks(relationships)
        @hyperlinks ||= extract_hyperlinks(relationships)
      end

      # Get the dimensions for the sheet.
      # This is the upper bound of cells that might
      # be parsed. (the document may be sparse so cell count is only upper bound)
      def dimensions
        @dimensions ||= extract_dimensions
      end

      # Yield each row xml element to caller
      def each_row_streaming(&block)
        Roo::Utils.each_element(@path, 'row', &block)
      end

      # Yield each cell as Excelx::Cell to caller for given
      # row xml
      def each_cell(row_xml)
        return [] unless row_xml
        row_xml.children.each do |cell_element|
          key = ::Roo::Utils.ref_to_key(cell_element['r'])
          yield cell_from_xml(cell_element, hyperlinks(@relationships)[key])
        end
      end

      private

      def cell_value_type(type, format)
        case type
        when 's'
          :shared
        when 'b'
          :boolean
        when 'str'
          :string
        when 'inlineStr'
          :inlinestr
        else
          Excelx::Format.to_type(format)
        end
      end

      # Internal:
      #
      # cell_xml - a Nokogiri::XML::Element. e.g.
      #             <c r="A5" s="2">
      #               <v>22606</v>
      #             </c>
      # hyperlink - a String for the hyperlink for the cell or nil when no
      #             hyperlink is present.
      #
      # Examples
      #
      #    cells_from_xml(<Nokogiri::XML::Element>, nil)
      #    # => <Excelx::Cell::String>
      #
      # Returns a type of <Excelx::Cell>.
      def cell_from_xml(cell_xml, hyperlink)
        coordinate = extract_coordinate(cell_xml['r'])
        return Excelx::Cell::Empty.new(coordinate) if cell_xml.children.empty?

        # NOTE: This is error prone, to_i will silently turn a nil into a 0.
        #       This works by coincidence because Format[0] is General.
        style = cell_xml['s'].to_i
        format = @styles.style_format(style)
        value_type = cell_value_type(cell_xml['t'], format)
        formula = nil

        cell_xml.children.each do |cell|
          case cell.name
          when 'is'
            cell.children.each do |inline_str|
              if inline_str.name == 't'
                return Excelx::Cell::String.new(inline_str.content, formula, :string, style, hyperlink, @workbook.base_date, coordinate)
              end
            end
          when 'f'
            formula = cell.content
          when 'v'
            return create_cell_from_value(value_type, cell, formula, format, style, hyperlink, @workbook.base_date, coordinate)
          end
        end
      end

      def create_cell_from_value(value_type, cell, formula, format, style, hyperlink, base_date, coordinate)
        # TODO: This can probably be removed. If that's the case
        excelx_type = [:numeric_or_formula, format.to_s]

        # TODO: cleanup and use to reate cells. it
        # NOTE: there are only a few situations where value != cell.content
        #       1. when a sharedString is used. value = sharedString;
        #          cell.content = id of sharedString
        #       2. boolean cells: value = 'TRUE' | 'FALSE'; cell.content = '0' | '1';
        #          In truth, I'd prefer a boolean cell that used TRUE|FALSE
        #          as the formatted value and an actual Boolean as the value.
        #       3. formula
        # FIXME: don't need base_date or excelx_type
        case value_type
        when :shared
          # FIXME: Doesn't need base_date
          value = @shared_strings[cell.content.to_i]
          excelx_type = :string
          return Excelx::Cell.create_cell(:string, value, formula, excelx_type, style, hyperlink, nil, coordinate)
        when :string
          # FIXME: Doesn't need base_date
          value = cell.content
          excelx_type = :string
          return Excelx::Cell.create_cell(value_type, value, formula, excelx_type, style, hyperlink, base_date, coordinate)
        when :boolean, :string
          # FIXME: Don't need base_date
          value = cell.content
          return Excelx::Cell.create_cell(value_type, value, formula, excelx_type, style, hyperlink, base_date, coordinate)
        when :time, :datetime
          cell_content = cell.content.to_f
          if cell_content < 1.0
            return Excelx::Cell::Time.new(cell.content, formula, excelx_type, style, hyperlink, base_date, coordinate)
          elsif  (cell_content - cell_content.floor).abs > 0.000001
            return Excelx::Cell::DateTime.new(cell.content, formula, excelx_type, style, hyperlink, base_date, coordinate)
          else
            return Excelx::Cell::Date.new(cell.content, formula, excelx_type, style, hyperlink, base_date, coordinate)
          end
        when :date
          return Excelx::Cell::Date.new(cell.content, formula, excelx_type, style, hyperlink, base_date, coordinate)
        else
          # FIXME: Doesn't need base_date
          return Excelx::Cell::Number.new(cell.content, formula, excelx_type, style, hyperlink, base_date, coordinate)
        end
      end

      def extract_coordinate(coordinate)
        row, column = ::Roo::Utils.split_coordinate(coordinate)

        Excelx::Coordinate.new(row, column)
      end

      def extract_hyperlinks(relationships)
        # FIXME: select the valid hyperlinks and then map those.
        Hash[doc.xpath('/worksheet/hyperlinks/hyperlink').map do |hyperlink|
          if hyperlink.attribute('id') && (relationship = relationships[hyperlink.attribute('id').text])
            [::Roo::Utils.ref_to_key(hyperlink.attributes['ref'].to_s), relationship.attribute('Target').text]
          end
        end.compact]
      end

      def expand_merged_ranges(cells)
        # Extract merged ranges from xml
        merges = {}
        doc.xpath('/worksheet/mergeCells/mergeCell').each do |mergecell_xml|
          tl, br = mergecell_xml['ref'].split(/:/).map { |ref| ::Roo::Utils.ref_to_key(ref) }
          for row in tl[0]..br[0] do
            for col in tl[1]..br[1] do
              next if row == tl[0] && col == tl[1]
              merges[[row, col]] = tl
            end
          end
        end
        # Duplicate value into all cells in merged range
        merges.each do |dst, src|
          cells[dst] = cells[src]
        end
      end

      def extract_cells(relationships)
        extracted_cells = Hash[doc.xpath('/worksheet/sheetData/row/c').map do |cell_xml|
          key = ::Roo::Utils.ref_to_key(cell_xml['r'])
          [key, cell_from_xml(cell_xml, hyperlinks(relationships)[key])]
        end]

        expand_merged_ranges(extracted_cells) if @options[:expand_merged_ranges]

        extracted_cells
      end

      def extract_dimensions
        Roo::Utils.each_element(@path, 'dimension') do |dimension|
          return dimension.attributes['ref'].value
        end
      end

=begin
Datei xl/comments1.xml
  <?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
  <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <authors>
      <author />
    </authors>
    <commentList>
      <comment ref="B4" authorId="0">
        <text>
          <r>
            <rPr>
              <sz val="10" />
              <rFont val="Arial" />
              <family val="2" />
            </rPr>
            <t>Kommentar fuer B4</t>
          </r>
        </text>
      </comment>
      <comment ref="B5" authorId="0">
        <text>
          <r>
            <rPr>
            <sz val="10" />
            <rFont val="Arial" />
            <family val="2" />
          </rPr>
          <t>Kommentar fuer B5</t>
        </r>
      </text>
    </comment>
  </commentList>
  </comments>
=end
=begin
    if @comments_doc[self.sheets.index(sheet)]
      read_comments(sheet)
    end
=end
    end
  end
end
