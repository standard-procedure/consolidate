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

      # Get the width of this image in EMU up to a maximum page width (also in EMU)
      def clamped_width(maximum = 7_772_400) = [width, maximum].min

      # Get the height of this image in EMU adjusted for a maximum page width (also in EMU)
      def clamped_height(maximum = 7_772_400) = (height * clamped_width(maximum).to_f / width.to_f).to_i

      def emu_per_width_pixel = 914_400 / dpi[:x]

      def emu_per_height_pixel = 914_400 / dpi[:y]
    end
  end
end
