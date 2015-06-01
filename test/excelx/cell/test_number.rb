require 'roo/excelx/cell/base'
require 'roo/excelx/cell/number'
require 'roo/link'

class TestRooExcelxCellNumber < Minitest::Test
  def number
    Roo::Excelx::Cell::Number
  end

  def test_float
    skip
  end

  def test_integer
    skip
  end

  def test_percent
    skip
  end

  # def test_formats
  #             'General'
  #             '0'
  #             '0.00'
  #             '#,##0'
  #             '#,##0.00'
  #             '0%'
  #             '0.00%'
  # end
end
