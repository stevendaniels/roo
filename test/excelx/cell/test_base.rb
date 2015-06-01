require 'date'
require 'roo/excelx/cell/base'
require 'roo/link'

# attr_reader :formula, :format, :hyperlink, :coordinate

class TestRooExcelxCellBase < Minitest::Test
  def base
    Roo::Excelx::Cell::Base
  end

  def base_date
    Date.new(1899, 12, 30)
  end

  def value
    'Hello World'
  end

  def test_cell_type_is_base
    cell = base.new(value, nil, [], nil, nil, base_date, nil)
    assert_equal :base, cell.type
  end

  def test_cell_type_is_formula_with_formula
    formula = true
    cell = base.new(value, formula, [], nil, nil, base_date, nil)
    assert_equal :formula, cell.type
  end

  def test_cell_type_is_link_with_formula
    hyperlink = true
    cell = base.new(value, nil, [], nil, hyperlink, base_date, nil)
    assert_equal :link, cell.type
  end

  def test_cell_value
    cell_value = value
    cell = base.new(cell_value, nil, [], nil, nil, base_date, nil)
    assert_equal cell_value, cell.cell_value
  end

  def test_not_empty?
    cell = base.new(value, nil, [], nil, nil, base_date, nil)
    refute cell.empty?
  end

  def test_cell_type_is_formula_with_formula
    formula = true
    cell = base.new(value, formula, [], nil, nil, base_date, nil)
    assert_equal :formula, cell.type
  end

  def test_cell_type_is_link_with_formula
    hyperlink = 'http://example.com'
    cell = base.new(value, nil, [], nil, hyperlink, base_date, nil)
    assert_equal :link, cell.type
  end

  def test_link?
    hyperlink = 'http://example.com'
    cell = base.new(value, nil, [], nil, hyperlink, base_date, nil)
    assert cell.link?
  end

  def test_formula?
    formula = true
    cell = base.new(value, formula, [], nil, nil, base_date, nil)
    assert cell.formula?
  end
end
