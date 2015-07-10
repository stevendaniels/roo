require 'date'
require 'roo/excelx/cell/base'
require 'roo/excelx/cell/datetime'
require 'roo/link'
require 'pry'

# TODO
# Look at formulas in excel - does not work with date/time
# attr_reader :cell_value, :formula, :format, :hyperlink, :coordinate

class TestRooExcelxCellDateTime < Minitest::Test
  def test_cell_value_is_datetime
    cell = datetime.new('30000.323212', nil, ['mm-dd-yy'], nil, nil, base_date, nil)
    assert_kind_of ::DateTime, cell.value
  end

  def test_cell_type_is_datetime
    cell = datetime.new('30000.323212', nil, [], nil, nil, base_date, nil)
    assert_equal :datetime, cell.type
  end

  def test_standard_formatted_value
    [
      ['mm-dd-yy', '01-25-15'],
      ['d-mmm-yy', '25-JAN-15'],
      ['d-mmm ', '25-JAN'],
      ['mmm-yy', 'JAN-15'],
      ['m/d/yy h:mm', '1/25/15 8:15']
    ].each do |format, formatted_value|
      cell = datetime.new '42029.34375', nil, [format], nil, nil, base_date, nil
      assert_equal formatted_value, cell.formatted_value
    end
  end

  def datetime
    Roo::Excelx::Cell::DateTime
  end

  def base_date
    Date.new(1899, 12, 30)
  end
end
