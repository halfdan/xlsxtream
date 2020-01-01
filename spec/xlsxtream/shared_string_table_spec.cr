require "spec"

# require "../../../xlsxtream/shared_string_table"

module Xlsxtream
  describe SharedStringTable do
    it "#references" do
      sst = SharedStringTable.new
      sst.references.should eq(0)
      sst["hello"]
      sst["hello"]
      sst["world"]
      sst.references.should eq(3)
    end

    it "#size" do
      sst = SharedStringTable.new
      sst.size.should eq(0)
      sst["hello"]
      sst["hello"]
      sst["world"]
      sst.size.should eq(2)
    end

    it "#value" do
      sst = SharedStringTable.new
      sst["hello"].should eq(0)
      sst["world"].should eq(1)
      sst["hello"].should eq(0)
    end
  end
end
