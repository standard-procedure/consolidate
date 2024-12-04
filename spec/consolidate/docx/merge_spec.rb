# frozen_string_literal: true

require "spec_helper"

RSpec.describe Consolidate::Docx::Merge do
  require "zip"

  before do
    FileUtils.rm_rf Dir.glob("tmp/*.docx")
  end

  describe "with a simple docx file" do
    let(:file_path) { "spec/files/mm.docx" }
    let(:data) { {name: "Alice Aadvark", company_name: "TinyCo", system_name: "Collabor8", package: "Corporate"} }
    let(:text_data) { data.select { |k, v| !k.to_s.end_with? "_image" } }
    let(:image_data) { data.select { |k, v| k.to_s.end_with? "_image" } }

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
        merge.write_to "tmp/simple.docx"
      end

      expect(File.exist?("tmp/simple.docx")).to be true

      zip = Zip::File.open("tmp/simple.docx")
      xml = zip.read("word/document.xml")
      text_data.values.each do |value|
        expect(xml.include?(value)).to eq true
      end
    end

    it "replaces the image fields with embedded images and references to those images" do
      Consolidate::Docx::Merge.open(file_path, verbose: true) do |merge|
        merge.data data
        merge.write_to "tmp/simple.docx"
      end

      expect(File.exist?("tmp/simple.docx")).to be true

      zip = Zip::File.open("tmp/simple.docx")
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
    let(:data) { {name: "Alice Aadvark", company_name: "TinyCo", system_name: "Collabor8", package: "Corporate", first_image: File.read(first_image_path), second_image: File.read(second_image_path)} }
    let(:text_data) { data.select { |k, v| !k.to_s.end_with? "_image" } }
    let(:image_data) { data.select { |k, v| k.to_s.end_with? "_image" } }
    let(:first_image_path) { "spec/files/c8o.png" }
    let(:second_image_path) { "spec/files/echodek.png" }

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
      expect(field_names).to eq(["first_image", "second_image"])
    end

    it "replaces the text fields with the supplied data" do
      Consolidate::Docx::Merge.open(file_path, verbose: false) do |merge|
        merge.data data
        merge.write_to "tmp/simple.docx"
      end

      expect(File.exist?("tmp/simple.docx")).to be true

      zip = Zip::File.open("tmp/simple.docx")
      xml = zip.read("word/document.xml")
      text_data.values.each do |value|
        expect(xml.include?(value)).to eq true
      end
    end

    it "replaces the image fields with embedded images and references to those images" do
      Consolidate::Docx::Merge.open(file_path, verbose: false) do |merge|
        merge.data data
        merge.write_to "tmp/simple.docx"
      end

      expect(File.exist?("tmp/simple.docx")).to be true

      zip = Zip::File.open("tmp/simple.docx")
      xml = zip.read("word/document.xml")
      image_data.each do |field_name, value|
        image_path = "word/media/#{field_name}.png"
        image_id = "rId#{field_name}"
        expect(xml.include?(image_id)).to eq true
        expect(zip.find_entry(image_path)).to be_truthy
      end
    end
  end

  describe "with a docx file where the text fields have embedded word formatting" do
    let(:file_path) { "spec/files/mangled.docx" }
    let(:data) { {"STN" => "0123 456789", "EML" => "alice@example.com"} }
    let(:text_data) { data.select { |k, v| !k.to_s.end_with? "_image" } }
    let(:image_data) { data.select { |k, v| k.to_s.end_with? "_image" } }

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
