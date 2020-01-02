module Xlsxtream
  enum FontFamily
    DEFAULT
    ROMAN
    SWISS
    MODERN
    SCRIPT
    DECORATIVE
  end

  class Font
    getter name
    getter size
    getter family

    def initialize(@name : String, @size : UInt8, @family : FontFamily = FontFamily::DEFAULT)
    end
  end
end
