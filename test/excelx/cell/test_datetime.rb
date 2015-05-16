require 'date'
require 'roo/excelx/cell/base'
require 'roo/excelx/cell/datetime'
require 'roo/link'

#TODO
# Look at formulas in excel - does not work with date/time
        # attr_reader :cell_value, :formula, :format, :hyperlink, :coordinate

class TestRooExcelxCellDateTime < Minitest::Test
  def test_cell_value_is_datetime
    cell = datetime.new('30000.323212', nil, [], nil, nil, base_date, nil)
    assert_kind_of ::DateTime, cell.value
  end

  def test_cell_type_is_datetime
    cell = datetime.new('30000.323212', nil, [], nil, nil, base_date, nil)
    assert_equal :datetime, cell.type
  end

  def test_cell_type_is_formula_with_formula
    # TODO: move to Cell::Base
    formula = true
    cell = datetime.new('30000.323212', formula, [], nil, nil, base_date, nil)
    assert_equal :formula, cell.type
  end

  def test_cell_type_is_link_with_formula
    # TODO: move to Cell::Base
    hyperlink = true
    cell = datetime.new('30000.323212', nil, [], nil, hyperlink, base_date, nil)
    assert_equal :link, cell.type
  end

  def datetime
    Roo::Excelx::Cell::DateTime
  end

  def base_date
    Date.new(1899, 12, 30)
  end
end

        # def type
        #   if formula?
        #     :formula
        #   elsif hyperlink?
        #     :link
        #   else
        #     @type
        #   end
        # end
        #
        # def formula?
        #   !!@formula
        # end
        #
        # def hyperlink?
        #   !!@hyperlink
        # end
        #
        # def formatted_value
        #   @value.strftime(@format.gsub(/#{formats.keys.join('|')}/, formats))
        # end
        #
        # # DEPRECATED: Please use cell_value instead.
        # def excelx_value
        #   warn '[DEPRECATION] `excelx_value` is deprecated.  Please use `cell_value` instead.'
        #   cell_value
        # end
        #
        # # DEPRECATED: Please use cell_type instead.
        # def excelx_type
        #   warn '[DEPRECATION] `excelx_type` is deprecated.  Please use `cell_type` instead.'
        #   cell_type
        # end
        #
        # # DEPRECATED: will be removed in next major version
        # def style
        #   warn '[DEPRECATION] `style` is deprecated and will be remove in next major version.'
        #   @style
        # end
