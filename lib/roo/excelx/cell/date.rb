require 'date'

module Roo
  class Excelx
    class Cell
      class Date < Roo::Excelx::Cell::DateTime
        attr_reader :value, :formula, :format, :cell_type, :cell_value, :link, :coordinate

        def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
          # NOTE: Pass all arguments to the parent class, DateTime.
          super
          @type = :date
          @format = excelx_type.last
          @value = link? ? Roo::Link.new(link, value) : create_date(base_date, value)
        end

        def formatted_value
          formatter = @format.downcase.gsub(/#{formats.keys.join('|')}/, formats)
          @value.strftime(formatter)
        end

        private

        def formats
          {
            'yyyy'.freeze => '%Y'.freeze,  # Year: 2000
            'yy'.freeze => '%y'.freeze,    # Year: 00
            'mmmm'.freeze => '%B'.freeze,  # Month: January
            'MMM'.freeze => '%b'.freeze,   # Month: Jan
            'mm'.freeze => '%m'.freeze,    # Month: 01
            'm'.freeze => '%-m'.freeze,    # Month: 1
            'dddd'.freeze => '%A'.freeze,  # Day of the Week: Sunday
            'ddd'.freeze => '%a'.freeze,   # Day of the Week: Sun
            'dd'.freeze => '%d'.freeze,    # Day of the Month: 01
            'd'.freeze => '%-d'.freeze,    # Day of the Month: 1
          }
        end

        def create_date(base_date, value)
          date = base_date + value.to_i
          yyyy, mm, dd = date.strftime('%Y-%m-%d').split('-')

          ::Date.new(yyyy.to_i, mm.to_i, dd.to_i)
        end
      end
    end
  end
end
