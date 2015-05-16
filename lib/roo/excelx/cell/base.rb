module Roo
  class Excelx
    class Cell
      class Base
        def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
          @link = !!link
          @cell_value = value
          @cell_type = excelx_type
          @formula = formula
          @style = style
          @coordinate = coordinate
        end

        def type
          if formula?
            :formula
          elsif link?
            :link
          else
            @type
          end
        end

        def formula?
          !!@formula
        end

        def link?
          !!@link
        end

        def formatted_value
          @value.strftime(@format.gsub(/#{formats.keys.join('|')}/, formats))
        end

        # DEPRECATED: Please use link instead.
        def hyperlink
          warn '[DEPRECATION] `hyperlink` is deprecated.  Please use `link` instead.'
        end

        # DEPRECATED: Please use cell_value instead.
        def excelx_value
          warn '[DEPRECATION] `excelx_value` is deprecated.  Please use `cell_value` instead.'
          cell_value
        end

        # DEPRECATED: Please use cell_type instead.
        def excelx_type
          warn '[DEPRECATION] `excelx_type` is deprecated.  Please use `cell_type` instead.'
          cell_type
        end

        # DEPRECATED: will be removed in next major version
        def style
          warn '[DEPRECATION] `style` is deprecated and will be remove in next major version.'
          @style
        end

        private

        def formats
          {
            'yyyy' => '%Y', # 2000
            'yy' => '%y',
            'mmmm' => '%B',
            'mmm' => '%b',
            'mm' => '%m',
            'm' => '%-m',
            'dddd' => '%A', # Sunday
            'ddd' => '%a', # Sun
            'dd' => '%d',
            'd' => '%-d',
            'hh' => '%h',
            'h' => '%-l',
            'ss' => '%s',
            'am/pm' => '%p'
          }
        end
      end
    end
  end
end
