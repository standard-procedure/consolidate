# frozen_string_literal: true

require "zip"
require "nokogiri"
require_relative "image_reference_node_builder"
require_relative "image"

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
        @contents_xml = load_and_update_contents_xml
        @output = {}
        @images = {}
        @mapping = {}
      end

      # Helper method to display the contents of the document and the merge fields from the CLI
      def examine
        puts "Documents: #{document_names.join(", ")}"
        puts "Content documents: #{content_document_names.join(", ")}"
        puts "Merge fields: #{text_field_names.join(", ")}"
        puts "Image fields: #{image_field_names.join(", ")}"
      end

      # Read all documents within the docx and extract any merge fields
      def text_field_names = @text_field_names ||= tag_nodes.collect { |tag_node| text_field_names_from tag_node }.flatten.compact.uniq

      # Read all documents within the docx and extract any image fields
      def image_field_names = @image_field_names ||= tag_nodes.collect { |tag_node| image_field_names_from tag_node }.flatten.compact.uniq

      # List the documents stored within this docx
      def document_names = @zip.entries.map(&:name)

      # List the content within this docx
      def content_document_names = @documents.keys

      # List the field names that are present in the merge data
      def merge_field_names = @mapping.keys

      # Set the merge data and erform the substitution - creating copies of any documents that contain merge tags and replacing the tags with the supplied data
      def data mapping = {}
        @mapping = mapping.transform_keys(&:to_s)
        if verbose
          puts "...mapping data"
          puts @mapping.keys.select { |field_name| text_field_names.include?(field_name) }.map { |field_name| "...   #{field_name} => #{@mapping[field_name]}" }.join("\n")
        end

        @images = load_images_and_link_relations

        @documents.each do |name, document|
          @output[name] = substitute(document.dup, document_name: name).serialize save_with: 0
        end
      end

      def write_to path
        puts "...writing to #{path}" if verbose
        Zip::File.open(path, create: true) do |out|
          @output[contents_xml] = @contents_xml.serialize save_with: 0

          @images.each do |field_name, image|
            next if image.nil?
            puts "...  writing image #{field_name} to #{image.storage_path}" if verbose
            out.get_output_stream(image.storage_path) { |o| o.write image.contents }
          end

          @relations.each do |relation_name, relations|
            puts "...  writing relations #{relation_name}" if verbose
            out.get_output_stream(relation_name) { |o| o.write relations }
          end

          @zip.reject do |entry|
            @relations.key? entry.name
          end.each do |entry|
            puts "...  writing updated document to #{entry.name}" if verbose
            out.get_output_stream(entry.name) { |o| o.write(@output[entry.name] || @relations[entry.name] || @zip.read(entry.name)) }
          end
        end
      end

      private

      attr_reader :verbose

      def contents_xml = "[Content_Types].xml"

      # Regex to find merge fields that contain text
      def text_tag = /\{\{\s*(?!.*_image\b)(\S+)\s*\}\}/i

      # Regex to find merge fields that contain images
      def image_tag = /\{\{\s*(\S+_image)\s*\}\}/i

      # Regex to find merge fields containing the given field name
      def tag_for(field_name) = /\{\{\s*#{field_name}\s*\}\}/

      # Find all nodes in all relevant documents that contain a merge field
      def tag_nodes = @documents.collect { |name, document| tag_nodes_for document }.flatten

      # go through all paragraph nodes of the document
      # selecting any that contain a merge tag
      def tag_nodes_for(document) = (document / "//w:p").select { |paragraph| paragraph.content.match(text_tag) || paragraph.content.match(image_tag) }

      # Extract the text field name(s) from the paragraph
      def text_field_names_from(tag_node) = (matches = tag_node.content.scan(text_tag)).empty? ? nil : matches.flatten.map(&:strip)

      # Extract the image field name(s) from the paragraph
      def image_field_names_from(tag_node) = (matches = tag_node.content.scan(image_tag)).empty? ? nil : matches.flatten.map(&:strip)

      # Unique number for each image field
      def relation_number_for(field_name) = @mapping.keys.index(field_name) + 1000

      # Identifier to use when linking a merge field to the actual image file contents
      def relation_id_for(field_name) = "rId#{field_name}"

      # Empty elations document for documents that do not already have one
      def default_relations_document = %(<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>)

      def load_documents
        @zip.entries.each_with_object({}) do |entry, results|
          next unless entry.name.match?(/word\/(document|header|footer|footnotes|endnotes).?\.xml/)
          puts "...reading document #{entry.name}" if verbose
          contents = @zip.get_input_stream entry
          results[entry.name] = Nokogiri::XML(contents) { |x| x.noent }
        end
      end

      def load_relations
        @zip.entries.each_with_object({}) do |entry, results|
          next unless entry.name.match?(/word\/(document|header|footer|footnotes|endnotes).?\.xml/)
          relation_document = entry.name.gsub("word/", "word/_rels/").gsub(".xml", ".xml.rels")
          puts "...reading or building relations for #{relation_document}" if verbose
          contents = @zip.find_entry(relation_document) ? @zip.get_input_stream(relation_document) : default_relations_document
          results[relation_document] = Nokogiri::XML(contents) { |x| x.noent }
        end
      ensure
        @zip.close
      end

      # Create relation links for each image field and store the image data
      def load_images_and_link_relations
        load_images.tap do |images|
          link_relations_to images
        end
      end

      # Build a mapping of image paths to the image data so that the image data can be stored in the output docx
      def load_images
        image_field_names.each_with_object({}) do |field_name, result|
          result[field_name] = @mapping[field_name].nil? ? nil : Consolidate::Docx::Image.new(@mapping[field_name])
          puts "...   #{field_name} => #{result[field_name]&.media_path}" if verbose
        end
      end

      # Update all relation documents to include a relationship for each image field and its stored image path
      def link_relations_to images
        @relations.each do |name, xml|
          puts "...   linking images in #{name}" if verbose
          images.each do |field_name, image|
            # Has an actual image file been supplied?
            next if image.nil? 
            # Is this image already referenced in this relationship document?
            next unless xml.at_xpath("//Relationship[@Target=\"#{image.media_path}\"]").nil?
            puts "...      #{relation_id_for(field_name)} => #{image.media_path}" if verbose
            xml.root << Nokogiri::XML::Node.new("Relationship", xml).tap do |relation|
              relation["Id"] = relation_id_for(field_name)
              relation["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
              relation["Target"] = image.media_path
            end
          end
        end
      end

      def load_and_update_contents_xml
        puts "...reading and updating #{contents_xml}" if verbose
        content = @zip.get_input_stream(contents_xml)
        Nokogiri::XML(content) { |x| x.noent }.tap do |document|
          add_content_relations_to document
        end
      end

      # Go through the given document, replacing any merge fields with the values provided
      # and storing the results in a new document
      def substitute document, document_name:
        puts "...substituting fields in #{document_name}" if verbose && tag_nodes_for(document).any?
        substitute_text document, document_name: document_name
        substitute_images document, document_name: document_name
      end

      def substitute_text document, document_name:
        tag_nodes_for(document).each do |tag_node|
          field_names = text_field_names_from(tag_node) || []

          # Extract the properties (formatting) nodes if they exist
          paragraph_properties = tag_node.search ".//w:pPr"
          run_properties = tag_node.at_xpath ".//w:rPr"

          # Get the current contents, then substitute any text fields
          text = tag_node.content

          field_names.each do |field_name|
            field_value = @mapping[field_name].to_s
            puts "...   substituting '#{field_name}' with '#{field_value}'" if verbose
            text = text.gsub(tag_for(field_name), field_value)
          end

          # Create a new text node with the substituted text
          text_node = Nokogiri::XML::Node.new("w:t", tag_node.document)
          text_node.content = text

          # Create a new run node to hold the run properties and substitute text node
          run_node = Nokogiri::XML::Node.new("w:r", tag_node.document)
          run_node << run_properties unless run_properties.nil?
          run_node << text_node
          # Add the paragraph properties and the run node to the tag node
          tag_node.children = Nokogiri::XML::NodeSet.new(document, paragraph_properties.to_a + [run_node])
        rescue => ex
          # Have to mangle the exception message otherwise it outputs the entire document
          puts ex.message.to_s[0..255]
          puts ex.backtrace.first
        end
        document
      end

      # Go through the given document, replacing any merge fields with the values provided
      # and storing the results in a new document
      def substitute_images document, document_name:
        tag_nodes_for(document).each do |tag_node|
          field_names = image_field_names_from(tag_node) || []
          # Extract the properties (formatting) nodes if they exist
          paragraph_properties = tag_node.search ".//w:pPr"
          run_properties = tag_node.at_xpath ".//w:rPr"

          pieces = tag_node.content.split(image_tag)
          # Split the content into pieces - either text or an image merge field
          # Then replace the text with text nodes or the image merge fields with drawing nodes
          replacement_nodes = pieces.collect do |piece|
            field_name = piece.strip
            if field_names.include? field_name
              image = @images[field_name]
              # if no image was provided then insert blank text
              # otherwise insert a w:drawing node that references the image contents
              if image.nil? 
                puts "...   substituting '#{field_name}' with blank as no image was provided" if verbose
                Nokogiri::XML::Node.new("w:t", document) { |t| t.content = "" }
              else
                puts "...   substituting '#{field_name}' with '<#{relation_id_for(field_name)}/>'" if verbose
                ImageReferenceNodeBuilder.new(field_name: field_name, image: image, node_id: relation_id_for(field_name), image_number: relation_number_for(field_name), document: document).call
              end
            else
              Nokogiri::XML::Node.new("w:t", document) { |t| t.content = piece }
            end
          end
          run_nodes = (replacement_nodes.map { |node| Nokogiri::XML::Node.new("w:r", document) { |run_node| run_node.children = node } } + [run_properties]).compact
          tag_node.children = Nokogiri::XML::NodeSet.new(document, paragraph_properties.to_a + run_nodes)
        rescue => ex
          # Have to mangle the exception message otherwise it outputs the entire document
          puts ex.message.to_s[0..255]
          puts ex.backtrace.first
        end
        document
      end

      CONTENT_RELATIONS = {
        jpeg: "image/jpg",
        png: "image/png",
        bmp: "image/bmp",
        gif: "image/gif",
        tif: "image/tif",
        pdf: "application/pdf",
        mov: "application/movie"
      }.freeze

      def add_content_relations_to document
        CONTENT_RELATIONS.each do |file_type, content_type|
          next unless document.at_xpath("//Default[@Extension=\"#{file_type}\"]").nil?
          document.root << Nokogiri::XML::Node.new("Default", document).tap do |relation|
            relation["Extension"] = file_type
            relation["ContentType"] = content_type
          end
        end
      end
    end
  end
end
