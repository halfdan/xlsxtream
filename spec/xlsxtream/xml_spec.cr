module Xlsxtream
  describe XML do
    it "returns a valid header" do
      expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      XML.header.should eq(expected)
    end

    it "strips newlines" do
      xml = <<-XML
      <hello id="1">
          <world/>
      </hello>
      XML
      expected = "<hello id=\"1\"><world/></hello>"
      XML.strip(xml).should eq(expected)
    end

    it "escapes attributes" do
      unsafe_attribute = "<hello> & \"world\""
      expected = "&lt;hello&gt; &amp; &quot;world&quot;"
      XML.escape_attr(unsafe_attribute).should eq(expected)
    end

    it "escapes XML reserved values" do
      unsafe_value = "<hello> & \"world\""
      expected = "&lt;hello&gt; &amp; \"world\""
      XML.escape_value(unsafe_value).should eq(expected)
    end

    it "escapes invalid XML characters" do
      unsafe_value = "The \x07 rings\x00\uFFFE\uFFFF"
      expected = "The _x0007_ rings_x0000__xFFFE__xFFFF_"
      XML.escape_value(unsafe_value).should eq(expected)
    end

    it "does not escape valid XML characters" do
      safe_value = "\u{10000}\u{10FFFF}"
      expected = safe_value 
      XML.escape_value(safe_value).should eq(expected)
    end
  end
end
