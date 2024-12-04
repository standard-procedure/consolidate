# frozen_string_literal: true

require "zip"
require "nokogiri"
require_relative "image_reference_node_builder"

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
        @zip = Zip::File.open(path)
        @documents = load_documents
        @relations = load_relations
        @output = {}
        @images = {}
      end

      # Helper method to display the contents of the document and the merge fields from the CLI
      def examine
        puts "Documents: #{document_names.join(", ")}"
        puts "Merge fields: #{text_field_names.join(", ")}"
        puts "Image fields: #{image_field_names.join(", ")}"
      end

      # Read all documents within the docx and extract any merge fields
      def text_field_names = @text_field_names ||= tag_nodes.collect { |tag_node| text_field_names_from tag_node }.flatten.compact.uniq

      # Read all documents within the docx and extract any image fields
      def image_field_names = @image_field_names ||= tag_nodes.collect { |tag_node| image_field_names_from tag_node }.flatten.compact.uniq

      # List the documents stored within this docx
      def document_names = @zip.entries.collect { |entry| entry.name }

      # Set the merge data and erform the substitution - creating copies of any documents that contain merge tags and replacing the tags with the supplied data
      def data mapping = {}
        mapping = mapping.transform_keys(&:to_s)
        puts mapping.keys.select { |field_name| text_field_names.include?(field_name) }.map { |field_name| "#{field_name} => #{mapping[field_name]}" }.join("\n") if verbose

        @images = load_images_and_link_relations_from mapping

        @documents.each do |name, document|
          @output[name] = substitute(document.dup, mapping: mapping, document_name: name).serialize save_with: 0
        end
      end

      def write_to path
        puts "...writing to #{path}" if verbose
        Zip::File.open(path, Zip::File::CREATE) do |out|
          @images.each do |image_name, data|
            puts "...  writing #{image_name} (#{data.size})" if verbose
            out.get_output_stream("word/#{image_name}") do |o|
              o.write data
            end
          end

          @zip.each do |entry|
            puts "...  writing #{entry.name}" if verbose
            out.get_output_stream(entry.name) do |o|
              o.write(@output[entry.name] || @relations[entry.name] || @zip.read(entry.name))
            end
          end
        end
      end

      private

      attr_reader :verbose

      def load_documents
        @zip.entries.each_with_object({}) do |entry, results|
          next unless entry.name.match?(/word\/(document|header|footer|footnotes|endnotes).?\.xml/)
          puts "...reading #{entry.name}" if verbose
          contents = @zip.get_input_stream entry
          results[entry.name] = Nokogiri::XML(contents) { |x| x.noent }
        end
      end

      def load_relations
        @zip.entries.each_with_object({}) do |entry, results|
          next unless entry.name.match?(/word\/_rels\/.*.rels/)
          puts "...reading #{entry.name}" if verbose
          contents = @zip.get_input_stream entry
          results[entry.name] = Nokogiri::XML(contents) { |x| x.noent }
        end
      ensure
        @zip.close
      end

      # Regex to find merge fields that contain text
      def text_tag = /\{\{\s*(?!.*_image\b)(\S+)\s*\}\}/i

      # Regex to find merge fields that contain images
      def image_tag = /\{\{\s*(\S+_image)\s*\}\}/i

      # Regex to find merge fields containing the given field name
      def tag_for(field_name) = /\{\{\s*#{field_name}\s*\}\}/

      # Path to store the image representing the given field name within the output docx
      def image_path_for(field_name) = "media/#{field_name}.png"

      # Find all nodes in all relevant documents that contain a merge field
      def tag_nodes = @documents.collect { |name, document| tag_nodes_for document }.flatten

      # go through all paragraph nodes of the document
      # selecting any that contain a merge tag
      def tag_nodes_for(document) = (document / "//w:p").select { |paragraph| paragraph.content.match(text_tag) || paragraph.content.match(image_tag) }

      # Extract the text field name(s) from the paragraph
      def text_field_names_from(tag_node) = (matches = tag_node.content.scan(text_tag)).empty? ? nil : matches.flatten.map(&:strip)

      # Extract the image field name(s) from the paragraph
      def image_field_names_from(tag_node) = (matches = tag_node.content.scan(image_tag)).empty? ? nil : matches.flatten.map(&:strip)

      # Go through the given document, replacing any merge fields with the values provided
      # and storing the results in a new document
      def substitute document, document_name:, mapping: {}
        tag_nodes_for(document).each do |tag_node|
          text_field_names = text_field_names_from(tag_node) || []
          image_field_names = image_field_names_from(tag_node) || []

          # Extract the properties (formatting) nodes if they exist
          paragraph_properties = tag_node.search ".//w:pPr"
          run_properties = tag_node.at_xpath ".//w:rPr"

          # Get the current contents, then substitute any text fields, followed by any image fields
          text = tag_node.content

          text_field_names.each do |field_name|
            field_value = mapping[field_name].to_s
            puts "...substituting #{field_name} with #{field_value} in #{document_name}" if verbose
            text = text.gsub(tag_for(field_name), field_value)
          end
          image_nodes = image_field_names.collect do |field_name|
            puts "...substituting #{field_name} with #{image_path_for(field_name)} in #{document_name}" if verbose
            # Add an image reference node and remove the merge tag
            text = text.gsub(tag_for(field_name), "")
            image_reference_node_for(field_name, document: document)
          end

          # Create a new text node with the substituted text
          text_node = Nokogiri::XML::Node.new("w:t", tag_node.document)
          text_node.content = text

          # Create a new run node to hold the run properties and substitute text node
          run_node = Nokogiri::XML::Node.new("w:r", tag_node.document)
          run_node << run_properties unless run_properties.nil?
          run_node << text_node
          # Add the paragraph properties and the run node to the tag node
          tag_node.children = Nokogiri::XML::NodeSet.new(document, paragraph_properties.to_a + [run_node] + image_nodes)
        rescue => ex
          # Have to mangle the exception message otherwise it outputs the entire document
          puts ex.message.to_s[0..255]
          puts ex.backtrace.first
        end
        document
      end

      def relation_id_for(field_name) = "rId#{field_name}"

      # Create relation links for each image field and store the image data
      def load_images_and_link_relations_from mapping
        link_relations_to_image_fields
        load_images_from mapping
      end

      # Build a mapping of image paths to the image data so that the image data can be stored in the output docx
      def load_images_from mapping = {}
        image_field_names.each_with_object({}) do |field_name, result|
          result[image_path_for(field_name)] = mapping[field_name]
        end
      end

      # Update all relation documents to include a relationship for each image field and its stored image path
      def link_relations_to_image_fields
        @relations.each do |name, xml|
          image_field_names.each do |field_name|
            next if xml.at_xpath("//Relationship[@Target='#{image_path_for(field_name)}']")
            puts "...linking #{field_name} to #{image_path_for(field_name)}" if verbose
            relation_node = Nokogiri::XML::Node.new("Relationship", xml)
            relation_node["Id"] = relation_id_for(field_name)
            relation_node["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            relation_node["Target"] = image_path_for(field_name)
            xml.root << relation_node
          end
        end
      end

      # Build an XML node to embed the image into the document
      def image_reference_node_for(field_name, document:) = ImageReferenceNodeBuilder.new(field_name: field_name, node_id: relation_id_for(field_name), document: document).call
    end
  end
end
