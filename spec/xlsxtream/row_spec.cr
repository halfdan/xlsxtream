module Xlsxtream
  describe Row do
    it "works with an empty column" do
      row = Row(Nil).new([nil], 1)

      row.to_xml.should eq("<row r=\"1\"></row>")
    end

    it "supports string columns" do
      row = Row(String).new(["hello"], 1)
      row.to_xml.should eq("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>hello</t></is></c></row>")
    end
  end
end
