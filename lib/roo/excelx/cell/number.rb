module Roo
  class Excelx
    class Cell
      class Number < Cell::Base
        attr_reader :value, :formula, :format, :cell_value, :link, :coordinate

        def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
          super
          # FIXME: change to number. This will break brittle tests.
          @type = :float
          @format = excelx_type.last
          @value = link? ? Roo::Link.new(link, value) : create_numeric(value)
        end

        def create_numeric(number)
          case @format
          when 'General', '0'
            number.include?('.') ? number.to_f : number.to_i
          when /%/
            number.to_f
          when /\.0/
            number.to_f
          end
        end

        def formatted_value
          formatter = formats[@format]
          if formatter.is_a? Proc
            @cell_value.call(formatter)
          else
            Kernel.format(@cell_value, formatter)
          end
        end

        def formats
          {
            'General'=> '%.0f',
            '0' => '%.0f',
            '0.00' => '%.2f',
            '#,##0' => proc do |number|
              format('%.0f', number).reverse.gsub(/(\d{3})(?=\d)/, '\\1,').reverse
            end,
            '#,##0.00' => proc do |number|
              format('%.2f', number).reverse.gsub(/(\d{3})(?=\d)/, '\\1,').reverse
            end,
            '0%' =>  proc do |number|
              format('%d%', number * 100)
            end,
            '0.00%' => proc do |number|
              format('%.2f%', number * 100)
            end,
            11 => '0.00E+00',
            '11' => '%.2E%',
            37 => '#,##0 ;(#,##0)',
            38 => '#,##0 ;[Red](#,##0)',
            39 => '#,##0.00;(#,##0.00)',
            40 => '#,##0.00;[Red](#,##0.00)',
            48 => '##0.0E+0',
            49 => '@'
          }
        end

      end
    end
  end
end
