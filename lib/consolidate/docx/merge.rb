# frozen_string_literal: true

require "zip"
require "nokogiri"

module Consolidate
  module Docx
    class Merge
      def self.open(path, force_settings: true, &block)
        new(path, force_settings: force_settings, &block)
      end

      def examine
        extract_field_names
      end

      def data fields = {}
        fields = fields.transform_keys(&:to_s)

        @documents.each do |name, document|
          result = document.dup
          result = substitute_style_one_with result, fields
          result = substitute_style_two_with result, fields

          @output[name] = result.serialize save_with: 0
        end
      end

      def write_to path
        Zip::File.open(path, Zip::File::CREATE) do |out|
          zip.each do |entry|
            out.get_output_stream(entry.name) do |o|
              o.write(output[entry.name] || zip.read(entry.name))
            end
          end
        end
      end

      protected

      attr_reader :zip
      attr_reader :xml
      attr_reader :documents
      attr_accessor :output

      def initialize(path, force_settings: true, &block)
        raise "No block given" unless block
        @output = {}
        @documents = {}
        set_standard_settings if force_settings
        begin
          @zip = Zip::File.open(path)
          ["word/document.xml", "word/header1.xml", "word/footer1.xml"].each do |document|
            next unless @zip.find_entry(document)
            xml = @zip.read document
            @documents[document] = Nokogiri::XML(xml) { |x| x.noent }
            yield self
          end
        ensure
          @zip.close
        end
      end

      def extract_field_names
        (extract_style_one + extract_style_two).uniq
      end

      def extract_style_one
        documents.collect do |name, document|
          (document / "//w:fldSimple").collect do |field|
            value = field.attributes["instr"].value.strip
            value.include?("MERGEFIELD") ? value.gsub("MERGEFIELD", "").strip : nil
          end.compact
        end.flatten
      end

      def extract_style_two
        documents.collect do |name, document|
          (document / "//w:instrText").collect do |instr|
            value = instr.inner_text
            value.include?("MERGEFIELD") ? value.gsub("MERGEFIELD", "").strip : nil
          end.compact
        end.flatten
      end

      def substitute_style_one_with document, fields
        # Word's first way of doing things
        (document / "//w:fldSimple").each do |field|
          if field.attributes["instr"].value =~ /MERGEFIELD (\S+)/
            text_node = (field / ".//w:t").first
            next unless text_node
            text_node.inner_html = fields[$1].to_s
          end
        end
        document
      end

      def substitute_style_two_with document, fields
        # Word's second way of doing things
        (document / "//w:instrText").each do |instr|
          if instr.inner_text =~ /MERGEFIELD (\S+)/
            text_node = instr.parent.next_sibling.next_sibling.xpath(".//w:t").first
            text_node ||= instr.parent.next_sibling.next_sibling.next_sibling.xpath(".//w:t").first
            next unless text_node
            text_node.inner_html = fields[$1].to_s
          end
        end
        document
      end

      def set_standard_settings
        output["word/settings.xml"] = %(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"><w:zoom w:percent="100"/></w:settings>)
      end

      def close
        zip.close
      end
    end
  end
end
