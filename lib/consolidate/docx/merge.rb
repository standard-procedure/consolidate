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

      def initialize(path, verbose: false, &block)
        @verbose = verbose
        @output = {}
        @zip = Zip::File.open(path)
        @documents = load_documents
        block&.call self
      end

      # Helper method to display the contents of the document and the merge fields from the CLI
      def examine
        documents = document_names.join(", ")
        fields = field_names.join(", ")
        puts "Documents: #{documents}"
        puts "Merge fields: #{fields}"
      end

      # Read all documents within the docx and extract any merge fields
      def field_names
        tag_nodes.collect do |tag_node|
          field_names_from tag_node
        end.flatten.compact.uniq
      end

      # List the documents stored within this docx
      def document_names
        @zip.entries.collect { |entry| entry.name }
      end

      # Substitute the data from the merge fields with the values provided
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

      # Write the new document to the given path
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

      private

      attr_reader :verbose
      attr_reader :zip
      attr_reader :xml
      attr_reader :documents
      attr_accessor :output
      TAG = /\{\{\s*(\S+)\s*\}\}/

      def load_documents
        @zip.entries.each_with_object({}) do |entry, documents|
          next unless entry.name.match?(/word\/(document|header|footer|footnotes|endnotes).?\.xml/)
          puts "...reading #{entry.name}" if verbose
          xml = @zip.get_input_stream entry
          documents[entry.name] = Nokogiri::XML(xml) { |x| x.noent }
        end
      ensure
        @zip.close
      end

      # Collect all the nodes that contain merge fields
      def tag_nodes
        documents.collect do |name, document|
          tag_nodes_for document
        end.flatten
      end

      # go through all w:t (Word Text???) nodes of the document
      # find any nodes that contain "{{"
      # then find the ancestor node that also includes the ending "}}"
      # This collection of nodes contains all the merge fields for this document
      def tag_nodes_for document
        (document / "//w:p").select do |paragraph|
          paragraph.content.match(TAG)
        end
      end

      # Extract the merge field name from the node
      def field_names_from(tag_node)
        matches = tag_node.content.scan(TAG)
        matches.empty? ? nil : matches.flatten.map(&:strip)
      end

      # Go through the given document, replacing any merge fields with the values provided
      # and storing the results in a new document
      def substitute document, document_name:, mapping: {}
        tag_nodes_for(document).each do |tag_node|
          field_names = field_names_from tag_node
          puts "Original Node for #{field_names} is #{tag_node}" if verbose

          # Extract the paragraph properties node if it exists
          paragraph_properties = tag_node.search ".//w:pPr"
          run_properties = tag_node.at_xpath ".//w:rPr"

          text = tag_node.content
          field_names.each do |field_name|
            field_value = mapping[field_name].to_s
            puts "...substituting #{field_name} with #{field_value} in #{document_name}" if verbose
            text = text.gsub(/{{\s*#{field_name}\s*}}/, field_value)
          end

          # Create a new text node with the substituted text
          text_node = Nokogiri::XML::Node.new("w:t", tag_node.document)
          text_node.content = text

          # Create a new run node to hold the substituted text and the paragraph properties
          run_node = Nokogiri::XML::Node.new("w:r", tag_node.document)
          run_node << run_properties if run_properties
          run_node << text_node
          tag_node.children = Nokogiri::XML::NodeSet.new(document, paragraph_properties.to_a + [run_node])

          puts "TAG NODE FOR #{field_names} IS #{tag_node}" if verbose
        rescue => ex
          # Have to mangle the exception message otherwise it outputs the entire document
          puts ex.message.to_s[0..255]
        end
        document
      end
    end
  end
end
