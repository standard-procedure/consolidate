# frozen_string_literal: true

require "zip"
require "nokogiri"

module Consolidate
  module Docx
    class Merge
      def self.open(path, verbose: false, &block)
        new(path, verbose: verbose, &block).tap do |merge|
          block&.call merge
        end
        path
      end

      def initialize(path, verbose: false)
        @verbose = verbose
        @output = {}
        @zip = Zip::File.open(path)
        @documents = load_documents
      end

      # Helper method to display the contents of the document and the merge fields from the CLI
      def examine
        puts "Documents: #{document_names.join(", ")}"
        puts "Merge fields: #{field_names.join(", ")}"
      end

      # Read all documents within the docx and extract any merge fields
      def field_names = tag_nodes.collect { |tag_node| field_names_from tag_node }.flatten.compact.uniq

      # List the documents stored within this docx
      def document_names = @zip.entries.collect { |entry| entry.name }

      # Perform the substitution - creating copies of any documents that contain merge tags and replacing the tags with the supplied data
      def data mapping = {}
        mapping = mapping.transform_keys(&:to_s)

        if verbose
          puts "...substitutions..."
          mapping.each do |key, value|
            puts "      #{key} => #{value}"
          end
        end

        @documents.each do |name, document|
          output_document = substitute document.dup, mapping: mapping, document_name: name
          @output[name] = output_document.serialize save_with: 0
        end
      end

      def write_to path
        puts "...writing to #{path}" if verbose
        Zip::File.open(path, Zip::File::CREATE) do |out|
          @zip.each do |entry|
            out.get_output_stream(entry.name) do |o|
              o.write(@output[entry.name] || @zip.read(entry.name))
            end
          end
        end
      end

      private

      attr_reader :verbose

      def any_tag = /\{\{\s*(\S+)\s*\}\}/

      def tag_for(field_name) = /\{\{\s*#{field_name}\s*\}\}/

      def load_documents
        @documents = @zip.entries.each_with_object({}) do |entry, results|
          next unless entry.name.match?(/word\/(document|header|footer|footnotes|endnotes).?\.xml/)
          puts "...reading #{entry.name}" if verbose
          xml = @zip.get_input_stream entry
          results[entry.name] = Nokogiri::XML(xml) { |x| x.noent }
        end
      ensure
        @zip.close
      end

      def tag_nodes = @documents.collect { |name, document| tag_nodes_for document }.flatten

      # go through all paragraph nodes of the document
      # selecting any that contain a merge tag
      def tag_nodes_for(document) = (document / "//w:p").select { |paragraph| paragraph.content.match(any_tag) }

      # Extract the merge field name(s) from the paragraph
      def field_names_from(tag_node) = (matches = tag_node.content.scan(any_tag)).empty? ? nil : matches.flatten.map(&:strip)

      # Go through the given document, replacing any merge fields with the values provided
      # and storing the results in a new document
      def substitute document, document_name:, mapping: {}
        tag_nodes_for(document).each do |tag_node|
          field_names = field_names_from tag_node

          # Extract the properties (formatting) nodes if they exist
          paragraph_properties = tag_node.search ".//w:pPr"
          run_properties = tag_node.at_xpath ".//w:rPr"

          text = tag_node.content
          field_names.each do |field_name|
            field_value = mapping[field_name].to_s
            puts "...substituting #{field_name} with #{field_value} in #{document_name}" if verbose
            text = text.gsub(tag_for(field_name), field_value)
          end

          # Create a new text node with the substituted text
          text_node = Nokogiri::XML::Node.new("w:t", tag_node.document)
          text_node.content = text

          # Create a new run node to hold the run properties and substitute text node
          run_node = Nokogiri::XML::Node.new("w:r", tag_node.document)
          run_node << run_properties if run_properties
          run_node << text_node

          # Add the paragraph properties and the run node to the tag node
          tag_node.children = Nokogiri::XML::NodeSet.new(document, paragraph_properties.to_a + [run_node])
        rescue => ex
          # Have to mangle the exception message otherwise it outputs the entire document
          puts ex.message.to_s[0..255]
        end
        document
      end
    end
  end
end
