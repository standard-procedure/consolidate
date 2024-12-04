# frozen_string_literal: true

require "zip"
require "nokogiri"

module Consolidate
  module Docx
    class Image < SimpleDelegator
      # Path to use when referencing this image from other documents
      def media_path = "media/#{name}"

      # Path to use when storing this image within the docx
      def storage_path = "word/#{media_path}"

      # Convert width from pixels to EMU
      def width = super * EMU_PER_PIXEL

      # Convert height from pixels to EMU
      def height = super * EMU_PER_PIXEL

      EMU_PER_PIXEL = (914400 / 72)
    end
  end
end
