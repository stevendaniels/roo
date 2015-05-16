require 'pry'
require 'date'
# TODO: add tests for formatted_value
# TODO: add test for formula?
# TODO: add tests for hyperlink?
# TODO: add tests for cell::DateTime that are independent of spreadsheets.
module Roo
  class Excelx
    class Cell
      class DateTime < Cell::Base
        attr_reader :value, :formula, :format, :cell_value, :link, :coordinate
        def initialize(value, formula, excelx_type, style, link, base_date, coordinates)
          super
          @type = :datetime
          @format = excelx_type.last
          @value = link? ? Roo::Link.new(link, value) : create_datetime(base_date, value)
        end

        private

        def create_datetime(base_date, value)
          date = base_date + value.to_f.round(6)
          datetime_string = date.strftime('%Y-%m-%d %H:%M:%S.%N')
          t = round_datetime(datetime_string)

          ::DateTime.civil(t.year, t.month, t.day, t.hour, t.min, t.sec)
        end

        def round_datetime(datetime_string)
          /(?<yyyy>\d+)-(?<mm>\d+)-(?<dd>\d+) (?<hh>\d+):(?<mi>\d+):(?<ss>\d+.\d+)/ =~ datetime_string

          Time.new(yyyy.to_i, mm.to_i, dd.to_i, hh.to_i, mi.to_i, ss.to_r).round(0)
        end
      end
    end
  end
end
