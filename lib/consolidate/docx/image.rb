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
      def width = super * emu_per_width_pixel

      # Convert height from pixels to EMU
      def height = super * emu_per_height_pixel

      def emu_per_width_pixel = EMU_PER_PIXEL * 72 / dpi[:x]

      def emu_per_height_pixel = EMU_PER_PIXEL * 72 / dpi[:y]

      DEFAULT_PAGE_WIDTH = 12_240
      TWENTIETHS_OF_A_POINT_TO_EMU = 635
      DEFAULT_PAGE_WIDTH_IN_EMU = DEFAULT_PAGE_WIDTH * TWENTIETHS_OF_A_POINT_TO_EMU
      EMU_PER_PIXEL = 9525
      DEFAULT_PAGE_HEIGHT = DEFAULT_PAGE_WIDTH * 11 / 8.5 # Assuming US Letter size
      DEFAULT_PAGE_HEIGHT_IN_EMU = DEFAULT_PAGE_HEIGHT * TWENTIETHS_OF_A_POINT_TO_EMU
    end
  end
end
