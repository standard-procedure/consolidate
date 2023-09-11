# frozen_string_literal: true

require "zip"
require "nokogiri"

module Consolidate
  module Docx
    class Merge
      def self.open(path, verbose: false, &block)
        new(path, verbose: verbose, &block)
        path
      end

      def examine
        puts "Documents: #{extract_document_names}"
        puts "Merge fields: #{extract_field_names}"
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

      attr_reader :verbose
      attr_reader :zip
      attr_reader :xml
      attr_reader :documents
      attr_accessor :output

      def initialize(path, verbose: false, &block)
        raise "No block given" unless block
        @verbose = verbose
        @output = {}
        @documents = {}
        begin
          @zip = Zip::File.open(path)
          @zip.entries.each do |entry|
            next unless entry.name =~ /word\/(document|header|footer|footnotes|endnotes).?\.xml/
            puts "Reading #{entry.name}" if verbose
            xml = @zip.get_input_stream entry
            @documents[entry.name] = Nokogiri::XML(xml) { |x| x.noent }
          end
          yield self
        ensure
          @zip.close
        end
      end

      def extract_document_names
        @zip.entries.collect { |entry| entry.name }.join(", ")
      end

      def extract_field_names
        (extract_style_one + extract_style_two).uniq.join(", ")
      end

      def extract_style_one
        documents.collect do |name, document|
          (document / "//w:fldSimple").collect do |field|
            value = field.attributes["instr"].value.strip
            puts "...found #{value} (v1) in #{name}" if verbose
            value.include?("MERGEFIELD") ? value.gsub("MERGEFIELD", "").strip : nil
          end.compact
        end.flatten
      end

      def extract_style_two
        documents.collect do |name, document|
          (document / "//w:instrText").collect do |instr|
            value = instr.inner_text
            puts "...found #{value} (v2) in #{name}" if verbose
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
            puts "...substituting v1 #{field.attributes["instr"]} with #{fields[$1]}" if verbose
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
            puts "...substituting v2 #{instr.inner_text} with #{fields[$1]}" if verbose
            text_node.inner_html = fields[$1].to_s
          end
        end
        document
      end

      def close
        zip.close
      end
    end
  end
end
