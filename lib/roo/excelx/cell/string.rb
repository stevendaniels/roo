module Roo
  class Excelx
    class Cell
      class String < Cell::Base
        attr_reader :value, :formula, :format, :cell_type, :cell_value, :link, :coordinate

        def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
          super
          @type = :string
          @format = excelx_type
          @value = link? ? Roo::Link.new(link, value) : value
        end

        alias_method :formatted_value, :value

        def empty?
          value.empty?
        end
      end
    end
  end
end
