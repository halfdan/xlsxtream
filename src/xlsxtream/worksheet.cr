require "./xml"
require "./shared_string_table"

module Xlsxtream
    class Worksheet
        def initialize(@io : IO, @sst : SharedStringTable?, @auto_format : Bool)
            @rownum = 1

            write_header
        end

        def close
            write_footer
        end

        def <<(row)
            @io << Row.new(row, @rownum, sst: @sst).to_xml
            @rownum += 1
        end

        private def write_header
            @io << XML.header
            @io << XML.strip(<<-XML)
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            XML

            # columns = Array(@options[:columns])
            # unless columns.empty?
            #     @io << Columns.new(columns).to_xml
            # end

            @io << XML.strip(<<-XML)
                <sheetData>
            XML
        end

        private def write_footer
            @io << XML.strip(<<-XML)
                </sheetData>
                </worksheet>
            XML
        end
    end
end