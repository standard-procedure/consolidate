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
          field_name_from tag_node
        end.compact.uniq
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
          fields.each do |key, value|
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
        (document / "//w:t").collect do |node|
          (node.children.any? { |child| child.content.include? "{{" }) ? enclosing_node_for_start_tag(node) : nil
        end.compact
      end

      # Extract the merge field name from the node
      def field_name_from(tag_node)
        return nil unless (matches = tag_node.content.match(/{{\s*(\S+)\s*}}/))
        field_name = matches[1].strip
        puts "...field #{field_name} found in #{name}" if verbose
        field_name.to_s
      end

      # Go through the given document, replacing any merge fields with the values provided
      # and storing the results in a new document
      def substitute document, document_name:, mapping: {}
        tag_nodes_for(document).each do |tag_node|
          field_name = field_name_from tag_node
          next unless mapping.has_key? field_name
          field_value = mapping[field_name]
          puts "...substituting #{field_name} with #{field_value} in #{document_name}" if verbose
          tag_node.content = tag_node.content.gsub(field_name, field_value).gsub(/{{\s*/, "").gsub(/\s*}}/, "")
        rescue => ex
          # Have to mangle the exception message otherwise it outputs the entire document
          puts ex.message.to_s[0..255]
        end
        document
      end

      # Find the ancestor node that contains both the start {{ text and the end }} text enclosing the merge field
      def enclosing_node_for_start_tag(node)
        return node if node.content.include? "}}"
        node.parent.nil? ? nil : enclosing_node_for_start_tag(node.parent)
      end
    end
  end
end
