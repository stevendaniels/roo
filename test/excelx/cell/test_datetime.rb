require 'date'
require 'roo/excelx/cell/base'
require 'roo/excelx/cell/datetime'
require 'roo/link'

# TODO
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

  def test_formatted_value
    skip
  end

  def datetime
    Roo::Excelx::Cell::DateTime
  end

  def base_date
    Date.new(1899, 12, 30)
  end
end
