module Roo
  class Excelx
    class Cell
      class Boolean < Cell::Base
        attr_reader :value, :formula, :format, :cell_type, :cell_value, :link, :coordinate

        def initialize(value, formula, excelx_type, style, link, base_date, coordinates)
          super
          @type = :boolean
          @format = excelx_type.last
          @value = link? ? Roo::Link.new(link, value) : create_boolean(value)
        end

        def formatted_value
          value == 1 ? 'TRUE' : 'FALSE'
        end

        private

        def create_boolean(value)
          # FIXME: Using a boolean will cause methods like Base#to_csv to fail.
          #       Roo is using some method to ignore false/nil values.
          value.to_i == 1 ? 1 : 0
        end
      end
    end
  end
end