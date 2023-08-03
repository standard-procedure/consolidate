# frozen_string_literal: true

require "spec_helper"

RSpec.describe Consolidate::Docx::Merge do
  require "zip"
  let(:file_path) { "spec/files/mm.docx" }
  let(:data) { {"Name" => "Alice Aadvark", "Company" => "TinyCo", "Package" => "Corporate"} }

  it "lists the merge fields within the document" do
    result = []
    Consolidate::Docx::Merge.open(file_path) do |merge|
      result = merge.examine
    end
    expect(result).to eq(["Name", "Company", "Package"])
  end

  it "performs a mailmerge and forces Word settings" do
    Consolidate::Docx::Merge.open(file_path) do |merge|
      merge.data data
      merge.write_to "tmp/output.docx"
    end

    expect(File.exist?("tmp/output.docx")).to be true

    zip = Zip::File.open("tmp/output.docx")
    xml = zip.read("word/document.xml")
    data.values.each do |value|
      expect(xml).to include(value)
    end
  end
end
