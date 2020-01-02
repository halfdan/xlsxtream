# frozen_string_literal: true
require "zip"

require "./xml"

require "./shared_string_table"
require "./worksheet"
require "./font"

module Xlsxtream
  class Workbook
    DEFAULT_FONT = Font.new("Calibri", 12, FontFamily::SWISS)

    def self.open(output : String, *, font = DEFAULT_FONT)
      workbook = new(output, font)
      begin
        yield workbook
      ensure
        workbook.close
      end
    end

    def self.open(output : String, *, font = DEFAULT_FONT) : Workbook
      new(output, font)
    end

    def initialize(@output : String, @font : Font)
      @file = File.open(output, "wb")
      @io = Zip::Writer.new(@file)
      @sst = SharedStringTable.new
      @worksheets = Hash(String, Int32).new ->(hash : Hash(String, Int32), key : String) { hash[key] = hash.size + 1 }
    end

    def close
      write_workbook
      write_styles
      write_sst unless @sst.empty?
      write_workbook_rels
      write_root_rels
      write_content_types
      @io.close
      @file.close if @file
      nil
    end

    def write_worksheet(name : String?, *, use_shared_strings : Bool = true, auto_format : Bool = false)
      name = name || "Sheet#{@worksheets.size + 1}"
      sst = use_shared_strings ? @sst : nil
      sheet_id = @worksheets[name]
      @io.add("xl/worksheets/sheet#{sheet_id}.xml") do |io|
        worksheet = Worksheet.new(io, sst, auto_format)
        yield worksheet
        worksheet.close
      end
      nil
    end

    private def write_workbook
      rid = "rId0"
      @io.add "xl/workbook.xml" do |io|
        io << XML.header
        io << XML.strip(<<-XML)
          <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <workbookPr date1904="false"/>
            <sheets>
        XML
        @worksheets.each do |name, sheet_id|
          rid = rid.succ
          io << "<sheet name=\"#{XML.escape_attr name}\" sheetId=\"#{sheet_id}\" r:id=\"#{rid}\"/>"
        end
        io << XML.strip(<<-XML)
            </sheets>
          </workbook>
        XML
      end
    end

    private def write_styles
      @io.add "xl/styles.xml" do |io|
        io << XML.header
        io << XML.strip(<<-XML)
          <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <numFmts count="2">
              <numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/>
              <numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd hh:mm:ss"/>
            </numFmts>
            <fonts count="1">
              <font>
                <sz val="#{XML.escape_attr @font.size.to_s}"/>
                <name val="#{XML.escape_attr @font.name}"/>
                <family val="#{@font.family.value}"/>
              </font>
            </fonts>
            <fills count="2">
              <fill>
                <patternFill patternType="none"/>
              </fill>
              <fill>
                <patternFill patternType="gray125"/>
              </fill>
            </fills>
            <borders count="1">
              <border/>
            </borders>
            <cellStyleXfs count="1">
              <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
            </cellStyleXfs>
            <cellXfs count="3">
              <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
              <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
              <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            </cellXfs>
            <cellStyles count="1">
              <cellStyle name="Normal" xfId="0" builtinId="0"/>
            </cellStyles>
            <dxfs count="0"/>
            <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
          </styleSheet>
        XML
      end
    end

    private def write_root_rels
      @io.add("_rels/.rels") do |io|
        io << XML.header
        io << XML.strip(<<-XML)
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
          </Relationships>
        XML
      end
    end

    private def write_sst
      @io.add("xl/sharedStrings.xml") do |io|
        io << XML.header
        io << "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"#{@sst.references}\" uniqueCount=\"#{@sst.size}\">"
        @sst.each_key do |string|
          io << "<si><t>#{XML.escape_value string}</t></si>"
        end
        io << "</sst>"
      end
    end

    private def write_workbook_rels
      rid = "rId0"
      @io.add("xl/_rels/workbook.xml.rels") do |io|
        io << XML.header
        io << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        @worksheets.each do |name, sheet_id|
          rid = rid.succ
          io << "<Relationship Id=\"#{rid}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet#{sheet_id}.xml\"/>"
        end
        rid = rid.succ
        io << "<Relationship Id=\"#{rid}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
        rid = rid.succ
        io << "<Relationship Id=\"#{rid}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>" unless @sst.empty?
        io << "</Relationships>"
      end
    end

    private def write_content_types
      @io.add("[Content_Types].xml") do |io|
        io << XML.header
        io << XML.strip(<<-XML)
          <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="xml" ContentType="application/xml"/>
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        XML
        io << "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" unless @sst.empty?
        @worksheets.each_value do |sheet_id|
          io << "<Override PartName=\"/xl/worksheets/sheet#{sheet_id}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        end
        io << "</Types>"
      end
    end
  end
end
