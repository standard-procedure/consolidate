# frozen_string_literal: true

require "zip"
require "nokogiri"

module Consolidate
  module Docx
    class ImageReferenceNodeBuilder < Data.define(:field_name, :image, :node_id, :image_number, :document)
      def call
        max_width, max_height = max_dimensions_from(document)
        scaled_width, scaled_height = scale_dimensions(image.width, image.height, max_width, max_height)

        Nokogiri::XML::Node.new("w:drawing", document).tap do |drawing|
          drawing["xmlns:a"] = "http://schemas.openxmlformats.org/drawingml/2006/main"
          drawing << Nokogiri::XML::Node.new("wp:inline", document).tap do |inline|
            inline["distT"] = "0"
            inline["distB"] = "0"
            inline["distL"] = "0"
            inline["distR"] = "0"
            inline << Nokogiri::XML::Node.new("wp:extent", document).tap do |extent|
              extent["cx"] = scaled_width
              extent["cy"] = scaled_height
            end
            inline << Nokogiri::XML::Node.new("wp:effectExtent", document).tap do |effect_extent|
              effect_extent["l"] = "0"
              effect_extent["t"] = "0"
              effect_extent["r"] = "0"
              effect_extent["b"] = "0"
            end
            inline << Nokogiri::XML::Node.new("wp:cNvGraphicFramePr", document).tap do |c_nv_graphic_frame_pr|
              c_nv_graphic_frame_pr << Nokogiri::XML::Node.new("a:graphicFrameLocks", document).tap do |graphic_frame_locks|
                graphic_frame_locks["noChangeAspect"] = true
              end
            end
            inline << Nokogiri::XML::Node.new("a:graphic", document).tap do |graphic|
              graphic["xmlns:a"] = "http://schemas.openxmlformats.org/drawingml/2006/main"
              graphic << Nokogiri::XML::Node.new("a:graphicData", document).tap do |graphic_data|
                graphic_data["uri"] = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                graphic_data << Nokogiri::XML::Node.new("pic:pic", document).tap do |pic|
                  pic["xmlns:pic"] = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                  pic << Nokogiri::XML::Node.new("pic:nvPicPr", document).tap do |nv_pic_pr|
                    nv_pic_pr << Nokogiri::XML::Node.new("pic:cNvPr", document).tap do |c_nv_pr|
                      c_nv_pr["id"] = image_number
                      c_nv_pr["name"] = image.name
                      c_nv_pr["descr"] = image.name
                      c_nv_pr["hidden"] = false
                      c_nv_pr << Nokogiri::XML::Node.new("pic:cNvPicPr", document)
                    end
                  end
                  pic << Nokogiri::XML::Node.new("pic:blipFill", document).tap do |blip_fill|
                    blip_fill << Nokogiri::XML::Node.new("a:blip", document).tap do |blip|
                      blip["r:embed"] = node_id
                    end
                    blip_fill << Nokogiri::XML::Node.new("a:stretch", document).tap do |stretch|
                      stretch << Nokogiri::XML::Node.new("a:fillRect", document)
                    end
                  end
                  pic << Nokogiri::XML::Node.new("pic:spPr", document).tap do |sp_pr|
                    sp_pr << Nokogiri::XML::Node.new("a:xfrm", document).tap do |xfrm|
                      xfrm << Nokogiri::XML::Node.new("a:off", document).tap do |off|
                        off["x"] = "0"
                        off["y"] = "0"
                      end
                      xfrm << Nokogiri::XML::Node.new("a:ext", document).tap do |ext|
                        ext["cx"] = scaled_width
                        ext["cy"] = scaled_height
                      end
                    end
                    sp_pr << Nokogiri::XML::Node.new("a:prstGeom", document).tap do |prst_geom|
                      prst_geom["prst"] = "rect"
                      prst_geom << Nokogiri::XML::Node.new("a:avLst", document)
                    end
                  end
                end
              end
            end
          end
        end
      end

      DEFAULT_PAGE_WIDTH = 12_240
      TWENTIETHS_OF_A_POINT_TO_EMU = 635
      DEFAULT_PAGE_WIDTH_IN_EMU = DEFAULT_PAGE_WIDTH * TWENTIETHS_OF_A_POINT_TO_EMU
      EMU_PER_PIXEL = 9525
      DEFAULT_PAGE_HEIGHT = DEFAULT_PAGE_WIDTH * 11 / 8.5 # Assuming standard page ratio
      DEFAULT_PAGE_HEIGHT_IN_EMU = DEFAULT_PAGE_HEIGHT * TWENTIETHS_OF_A_POINT_TO_EMU

      private def max_width_from document
        page_width = (document.at_xpath("//w:sectPr/w:pgSz/@w:w")&.value || DEFAULT_PAGE_WIDTH).to_i
        page_width * TWENTIETHS_OF_A_POINT_TO_EMU
      end

      private def max_dimensions_from(document)
        page_width = (document.at_xpath("//w:sectPr/w:pgSz/@w:w")&.value || DEFAULT_PAGE_WIDTH).to_i
        page_height = (document.at_xpath("//w:sectPr/w:pgSz/@w:h")&.value || DEFAULT_PAGE_HEIGHT).to_i

        width_emu = page_width * TWENTIETHS_OF_A_POINT_TO_EMU
        height_emu = page_height * TWENTIETHS_OF_A_POINT_TO_EMU

        [width_emu, height_emu]
      end

      private def scale_dimensions(width, height, max_width, max_height)
        width_ratio = max_width.to_f / width
        height_ratio = max_height.to_f / height
        scale = [width_ratio, height_ratio, 1.0].min # Never scale up

        [(width * scale).to_i, (height * scale).to_i]
      end
    end
  end
end
