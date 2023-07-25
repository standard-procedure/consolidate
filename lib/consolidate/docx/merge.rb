# frozen_string_literal: true

require "zip"
require "nokogiri"

module Consolidate
  module Docx
    class Merge
      def self.open(path, force_settings: true, &block)
        new(path, force_settings: force_settings, &block)
      end

      def data fields = {}
        fields = fields.transform_keys(&:to_s)

        xml = @zip.read("word/document.xml")
        doc = Nokogiri::XML(xml) { |x| x.noent }

        doc = substitute_style_one_with xml, doc, fields
        doc = substitute_style_two_with xml, doc, fields

        @output["word/document.xml"] = doc.serialize save_with: 0
      end

      def write_to path
        Zip::File.open(path, Zip::File::CREATE) do |out|
          @zip.each do |entry|
            out.get_output_stream(entry.name) do |o|
              o.write(@output[entry.name] || @zip.read(entry.name))
            end
          end
        end
      end

      protected

      def initialize(path, force_settings: true, &block)
        raise "No block given" unless block
        @output = {}
        set_standard_settings if force_settings
        begin
          @zip = Zip::File.open(path)
          yield self
        ensure
          @zip.close
        end
      end

      def substitute_style_one_with xml, doc, fields
        # Word's first way of doing things
        (doc / "//w:fldSimple").each do |field|
          if field.attributes["instr"].value =~ /MERGEFIELD (\S+)/
            text_node = (field / ".//w:t").first
            next unless text_node
            text_node.inner_html = fields[$1].to_s
          end
        end
        doc
      end

      def substitute_style_two_with xml, doc, fields
        # Word's second way of doing things
        (doc / "//w:instrText").each do |instr|
          if instr.inner_text =~ /MERGEFIELD (\S+)/
            text_node = instr.parent.next_sibling.next_sibling.xpath(".//w:t").first
            text_node ||= instr.parent.next_sibling.next_sibling.next_sibling.xpath(".//w:t").first
            next unless text_node
            text_node.inner_html = fields[$1].to_s
          end
        end
        doc
      end

      def set_standard_settings
        @output["word/settings.xml"] = %(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"><w:zoom w:percent="100"/></w:settings>)
      end

      def close
        @zip.close
      end
    end
  end
end
