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
        documents = document_names.join(", ")
        fields = field_names.join(", ")
        puts "Documents: #{documents}"
        puts "Merge fields: #{fields}"
      end

      def field_names
        documents.collect do |name, document|
          (document / "//w:t").collect do |text_node|
            next unless (matches = text_node.content.match(/{{\s*(\S+)\s*}}/))
            field_name = matches[1].strip
            puts "...field #{field_name} found in #{name}" if verbose
            field_name
          end.compact
        end.flatten
      end

      def document_names
        @zip.entries.collect { |entry| entry.name }
      end

      def data fields = {}
        fields = fields.transform_keys(&:to_s)

        if verbose
          puts "...substitutions..."
          fields.each do |key, value|
            puts "      #{key} => #{value}"
          end
        end

        @documents.each do |name, document|
          result = document.dup
          result = substitute result, fields, name

          @output[name] = result.serialize save_with: 0
        end
      end

      def write_to path
        puts "...writing to #{path}" if verbose
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
            next unless entry.name.match?(/word\/(document|header|footer|footnotes|endnotes).?\.xml/)
            puts "...reading #{entry.name}" if verbose
            xml = @zip.get_input_stream entry
            @documents[entry.name] = Nokogiri::XML(xml) { |x| x.noent }
          end
          yield self
        ensure
          @zip.close
        end
      end

      def substitute document, fields, document_name
        (document / "//w:t").each do |text_node|
          next unless (matches = text_node.content.match(/{{\s*(\S+)\s*}}/))
          field_name = matches[1].strip
          if fields.has_key? field_name
            field_value = fields[field_name]
            puts "...substituting #{field_name} with #{field_value} in #{document_name}" if verbose
            text_node.content = text_node.content.gsub(matches[1], field_value).gsub("{{", "").gsub("}}", "")
          elsif verbose
            puts "...found #{field_name} but no replacement value"
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
