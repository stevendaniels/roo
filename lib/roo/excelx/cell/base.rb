module Roo
  class Excelx
    class Cell
      class Base
        # TODO: cells can update their value. That messes up everything you are trying to do. setting a new value means
        #       creating a new cell. Not a fan.

        # NOTE: The main advantage to splitting Cell into cell_type classes is to
        #       allow each class to better format it's values.
        #
        #       Previously cell values were calculated by `Cell#type_cast_value`,
        #       but the methods did not always cast a value as expected (e.g.
        #       Integers were cast as Floats).
        #
        #       Additionally, cells have an `excelx_value` method which contained
        #       a string representation of the original value of the cell.
        #
        #       The new cell type classes keeps excelx_value as an alias to
        #       cell_value. Cell value is intended to be a method that can be
        #       used regardless of what kind of spreadsheet is being used. For
        #       cells using SharedStrings, an `excelx_value` now returns the
        #       string (it previously returns the index for the SharedString).
        #
        #       The new cell classes will also add a `formatted_value` method.
        #       This method returns a String representation of the display value
        #       seen in a spreadsheet program. These formatted values should
        #       probably be used when exporting to csv.
        #
        #       Finally, the original `value` method will be upgraded to better
        #       match the type of the value. Cell::Numbers will return an
        #       Integer or Float depending what type of number the value should
        #       be. Cell::Booleans will return a Boolean. (One exception:
        #       Cell::Time will continue to return seconds since midnight.)
        attr_reader :cell_type, :cell_value
        attr_writer :value

        def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
          @link = !!link
          @cell_value = value
          @cell_type = excelx_type
          @formula = formula
          @style = style
          @coordinate = coordinate
          @type = :base
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
          formatter = @format.gsub(/#{formats.keys.join('|')}/, formats)
          puts formatter
          @value.strftime(formatter)
        end

        alias_method :to_s, :formatted_value

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

        # NOTE: having a style attribute for a cell doesn't really accomplish
        #       much. Unless you want to export to excelx.
        def style
          @style
        end

        def empty?
          false
        end

        private

        def formats
          {
            # FIXME: missing formats for AD/BC and milliseconds.
            'yyyy' => '%Y', # 2000
            'yy' => '%y',
            'MMMM' => '%B',
            'MMM' => '%b',
            'MM' => '%m',
            'M' => '%-m',
            'dddd' => '%A', # Sunday
            'ddd' => '%a', # Sun
            'dd' => '%d',
            'd' => '%-d',
            'hh' => '%H',
            'h' => '%-l',
            'mm' => '%M',
            'ss' => '%S',
            'am/pm' => '%p',
            '\\\\' => '', # NOTE: Fixes output for custom formats.
          }
        end
      end
    end
  end
end
