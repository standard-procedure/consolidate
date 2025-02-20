# frozen_string_literal: true

require "spec_helper"

RSpec.describe Consolidate::Docx::Merge do
  require "zip"

  before :all do
    FileUtils.rm_rf Dir.glob("tmp/*.docx")
  end

  describe "with a simple docx file" do
    let(:file_path) { "spec/files/mm.docx" }
    let(:data) { {name: "Alice Aadvark", company_name: "TinyCo", system_name: "Collabor8", package: "Corporate"} }
    let(:text_data) { data.select { |k, v| !k.to_s.end_with? "_image" } }
    let(:image_data) { data.select { |k, v| k.to_s.end_with? "_image" } }
    let(:output_path) { "tmp/simple.docx" }

    it "lists the text fields within the document" do
      field_names = []
      Consolidate::Docx::Merge.open(file_path) do |merge|
        field_names = merge.text_field_names
      end
      expect(field_names).to eq(["name", "company_name", "system_name", "package"])
    end

    it "lists the image fields within the document" do
      field_names = []
      Consolidate::Docx::Merge.open(file_path) do |merge|
        field_names = merge.image_field_names
      end
      expect(field_names).to eq([])
    end

    it "replaces the text fields with the supplied data" do
      Consolidate::Docx::Merge.open(file_path, verbose: false) do |merge|
        merge.data data
        merge.write_to output_path
      end

      expect(File.exist?(output_path)).to be true

      zip = Zip::File.open(output_path)
      xml = zip.read("word/document.xml")
      text_data.values.each do |value|
        expect(xml.include?(value)).to eq true
      end
    end

    it "replaces the image fields with embedded images and references to those images" do
      Consolidate::Docx::Merge.open(file_path, verbose: false) do |merge|
        merge.data data
        merge.write_to output_path
      end

      expect(File.exist?(output_path)).to be true

      zip = Zip::File.open(output_path)
      xml = zip.read("word/document.xml")
      image_data.values.each do |value|
        image_path = "word/media/#{value}.png"
        expect(xml.include?(image_path)).to eq true
        expect(zip.find_entry(image_path)).to_not be_nil
      end
    end
  end

  describe "embedding images into the docx file" do
    let(:file_path) { "spec/files/embed.docx" }
    let(:data) { {name: "Alice Aadvark", company_name: "TinyCo", system_name: "Collabor8", package: "Corporate", first_image: first_image, second_image: second_image} }
    let(:text_data) { data.select { |k, v| !k.to_s.end_with? "_image" } }
    let(:image_data) { data.select { |k, v| k.to_s.end_with? "_image" } }
    let(:first_image_path) { "spec/files/c8o.png" }
    let(:first_image) { Consolidate::Image.new(name: "c8o.png", width: 256, height: 61, path: first_image_path) }
    let(:second_image_path) { "spec/files/echodek.png" }
    let(:second_image) { Consolidate::Image.new(name: "echodek.png", width: 128, height: 128, contents: File.read(second_image_path)) }
    let(:output_path) { "tmp/images.docx" }

    it "lists the text fields within the document" do
      field_names = []
      Consolidate::Docx::Merge.open(file_path) do |merge|
        field_names = merge.text_field_names
      end
      expect(field_names).to eq(["name", "company_name", "system_name", "package"])
    end

    it "lists the image fields within the document" do
      field_names = []
      Consolidate::Docx::Merge.open(file_path) do |merge|
        field_names = merge.image_field_names
      end
      expect(field_names.size).to eq 2
      expect(field_names).to include "first_image"
      expect(field_names).to include "second_image"
    end

    it "replaces the text fields with the supplied data" do
      Consolidate::Docx::Merge.open(file_path, verbose: false) do |merge|
        merge.data data
        merge.write_to output_path
      end

      expect(File.exist?(output_path)).to be true

      zip = Zip::File.open(output_path)
      xml = zip.read("word/document.xml")
      text_data.values.each do |value|
        expect(xml.include?(value)).to eq true
      end
    end

    it "replaces the image fields with embedded images and references to those images" do
      first_image_id = second_image_id = ""

      Consolidate::Docx::Merge.open(file_path, verbose: false) do |merge|
        merge.data data
        first_image_id = merge.send :relation_id_for, "first_image"
        second_image_id = merge.send :relation_id_for, "second_image"
        merge.write_to output_path
      end

      expect(File.exist?(output_path)).to be true

      zip = Zip::File.open(output_path)

      # First Image is substituted in the header
      xml = zip.read("word/header1.xml")
      expect(xml.include?(first_image_id)).to eq true
      expect(zip.find_entry("word/media/c8o.png")).to_not be_nil

      # Second Image is substituted in the main document
      xml = zip.read("word/document.xml")
      expect(xml.include?(second_image_id)).to eq true
      expect(zip.find_entry("word/media/echodek.png")).to_not be_nil
    end

    it "replaces the image fields with blanks if the image itself is not provided" do
      output_path = "tmp/images_blank.docx"
      data[:second_image] = nil
      first_image_id = second_image_id = ""

      Consolidate::Docx::Merge.open(file_path, verbose: false) do |merge|
        merge.data data
        first_image_id = merge.send :relation_id_for, "first_image"
        second_image_id = merge.send :relation_id_for, "second_image"
        merge.write_to output_path
      end

      expect(File.exist?(output_path)).to be true

      zip = Zip::File.open(output_path)

      # First Image is substituted in the header
      xml = zip.read("word/header1.xml")
      expect(xml.include?(first_image_id)).to eq true
      expect(zip.find_entry("word/media/c8o.png")).to_not be_nil

      # Second Image is substituted in the main document
      xml = zip.read("word/document.xml")
      expect(xml.include?(second_image_id)).to_not eq true
      expect(xml.include?("{{second_image}}")).to_not eq true
      expect(zip.find_entry("word/media/echodek.png")).to be_nil
    end
  end

  describe "with a docx file where the text fields have embedded word formatting" do
    let(:file_path) { "spec/files/mangled.docx" }
    let(:data) { {"STN" => "0123 456789", "EML" => "alice@example.com", "PID" => "Now", "PRD" => "Then", "PAB" => "never"} }
    let(:text_data) { data.select { |k, v| !k.to_s.end_with? "_image" } }
    let(:image_data) { data.select { |k, v| k.to_s.end_with? "_image" } }
    let(:output_path) { "tmp/formatted.docx" }

    it "lists the text fields within the document" do
      field_names = []
      Consolidate::Docx::Merge.open(file_path) do |merge|
        field_names = merge.text_field_names
      end
      expect(field_names).to include "STN"
      expect(field_names).to include "EML"
    end

    it "lists the image fields within the document" do
      field_names = []
      Consolidate::Docx::Merge.open(file_path) do |merge|
        field_names = merge.image_field_names
      end
      expect(field_names).to eq([])
    end

    it "replaces the text fields with the supplied data" do
      Consolidate::Docx::Merge.open(file_path, verbose: false) do |merge|
        merge.data data
        merge.write_to "tmp/mangled.docx"
      end

      expect(File.exist?("tmp/mangled.docx")).to be true

      zip = Zip::File.open("tmp/mangled.docx")
      xml = zip.read("word/document.xml")
      text_data.values.each do |value|
        expect(xml.include?(value)).to eq true
      end
    end
  end
end
