module Xlsxtream
  class SharedStringTable < Hash(String, Int32)
    def initialize
      @references = 0
      super ->(hash : Hash(String, Int32), key : String) { hash[key] = hash.size }
    end

    def [](string)
      @references += 1
      super
    end

    def references
      @references
    end
  end
end
