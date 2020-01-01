# frozen_string_literal: true
module Xlsxtream
  module XML
    XML_ESCAPES = {
      "&"  => "&amp;",
      "\"" => "&quot;",
      "<"  => "&lt;",
      ">"  => "&gt;",
    }

    XML_DECLARATION = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"

    WS_AROUND_TAGS = /(?<=>)\s+|\s+(?=<)/

    UNSAFE_ATTR_CHARS  = /[&"<>]/
    UNSAFE_VALUE_CHARS = /[&<>]/

    # http://www.w3.org/TR/REC-xml/#NT-Char:
    # Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
    INVALID_XML10_CHARS = /[^\x09\x0A\x0D\x20-\x{D7FF}\x{E000}-\x{FFFD}\x{10000}-\x{10FFFF}]/

    # ST_Xstring escaping
    ESCAPE_CHAR = ->(c : Char) { "_x%04X_" % c.ord }

    def self.header
      XML_DECLARATION
    end

    def self.strip(xml)
      xml.gsub(WS_AROUND_TAGS, "")
    end

    def self.escape_attr(string)
      string.gsub(UNSAFE_ATTR_CHARS, XML_ESCAPES)
    end

    def self.escape_value(string)
      string.gsub(UNSAFE_VALUE_CHARS, XML_ESCAPES).gsub(INVALID_XML10_CHARS) do |c|
        "_x%04X_" % c[0].ord
      end
    end
  end
end
