module Xlsxtream
  class Row(T)
    ENCODING = "UTF-8"

    NUMBER_PATTERN = /\A-?[0-9]+(\.[0-9]+)?\z/
    # ISO 8601 yyyy-mm-dd
    DATE_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z/
    # ISO 8601 yyyy-mm-ddThh:mm:ss(.s)(Z|+hh:mm|-hh:mm)
    TIME_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,9})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})?\z/

    TRUE_STRING  = "true"
    FALSE_STRING = "false"

    DATE_STYLE = 1
    TIME_STYLE = 2

    def initialize(@row : Array(T), @rownum : Int32, *, @sst : SharedStringTable? = nil, @auto_format : Bool = false)
    end

    def to_xml
      column = "A"
      xml = "<row r=\"#{@rownum}\">"

      @row.each do |value|
        cid = "#{column}#{@rownum}"
        column = column.succ

        if @auto_format && value.is_a?(String)
          value = auto_format(value)
        end

        case value
        when Number
          xml += "<c r=\"#{cid}\" t=\"n\"><v>#{value}</v></c>"
        when Bool
          xml += "<c r=\"#{cid}\" t=\"b\"><v>#{value ? 1 : 0}</v></c>"
        when Time
          xml += "<c r=\"#{cid}\" s=\"#{TIME_STYLE}\"><v>#{time_to_oa_date(value)}</v></c>"
          # when Date
          #   xml += "<c r=\"#{cid}\" s=\"#{DATE_STYLE}\"><v>#{time_to_oa_date(value)}</v></c>"
        else
          value = value.to_s

          unless value.empty? # no xml output for for empty strings
            # value = value.encode(ENCODING) unless value.valid_encoding?

            if sst = @sst
              xml += "<c r=\"#{cid}\" t=\"s\"><v>#{sst[value]}</v></c>"
            else
              xml += "<c r=\"#{cid}\" t=\"inlineStr\"><is><t>#{XML.escape_value(value)}</t></is></c>"
            end
          end
        end
      end

      xml += "</row>"
    end

    # Detects and casts numbers, date, time in text
    private def auto_format(value)
      case value
      when TRUE_STRING
        true
      when FALSE_STRING
        false
      when NUMBER_PATTERN
        value.includes?('.') ? value.to_f : value.to_i
      when DATE_PATTERN
        # Date.parse(value) rescue value
      when TIME_PATTERN
        Time.parse!(value, "%Y-%m-%d %H:%M:%S %z") rescue value
      else
        value
      end
    end

    def time_to_oa_date(time)
      # Local dates are stored as UTC by truncating the offset:
      # 1970-01-01 00:00:00 +0200 => 1970-01-01 00:00:00 UTC
      # This is done because SpreadsheetML is not timezone aware.
      time.to_utc.to_unix_f / 24 / 3600 + 25569
    end
  end
end
