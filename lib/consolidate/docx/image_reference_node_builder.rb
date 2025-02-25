# frozen_string_literal: true

require "zip"
require "nokogiri"
# To test image scaling manually, from the console:
# Consolidate::Docx::Merge.open "spec/files/logo-doc.docx" do |doc|
#   doc.data logo_image: Consolidate::Image.new(name: "logo.png", width: 4096, height: 1122, path: "spec/files/logo.png")
#   doc.write_to "tmp/merge.docx"
# end
module Consolidate
  module Docx
    class ImageReferenceNodeBuilder < Data.define(:field_name, :image, :node_id, :image_number, :document)
      def call
        usable_width, usable_height = usable_dimensions_from(document)
        scaled_width, scaled_height = scale_dimensions(image.width_in_emu, image.height_in_emu, usable_width, usable_height)

        Nokogiri::XML::Node.new("w:drawing", document).tap do |drawing|
          drawing["xmlns:a"] = "http://schemas.openxmlformats.org/drawingml/2006/main"
          drawing << create_anchor_node(scaled_width, scaled_height)
        end
      rescue => ex
        puts ex.backtrace
      end

      private def create_anchor_node(scaled_width, scaled_height)
        Nokogiri::XML::Node.new("wp:anchor", document).tap do |anchor|
          # Basic anchor properties
          anchor["behindDoc"] = "0"
          anchor["distT"] = "0"
          anchor["distB"] = "0"
          anchor["distL"] = "0"
          anchor["distR"] = "0"
          anchor["simplePos"] = "0"
          anchor["locked"] = "0"
          anchor["layoutInCell"] = "1"
          anchor["allowOverlap"] = "1"
          anchor["relativeHeight"] = "2"

          # Position nodes
          anchor.children = create_position_nodes

          # Size node
          anchor << Nokogiri::XML::Node.new("wp:extent", document).tap do |extent|
            extent["cx"] = scaled_width
            extent["cy"] = scaled_height
          end

          # Effect extent
          anchor << create_effect_extent_node

          # Wrap node
          anchor << Nokogiri::XML::Node.new("wp:wrapSquare", document).tap do |wrap_square|
            wrap_square["wrapText"] = "bothSides"
          end

          # Document properties
          anchor << create_doc_properties_node

          # Non-visual properties
          anchor << create_non_visual_properties_node

          # Graphic node
          anchor << create_graphic_node(scaled_width, scaled_height)
        end
      end

      private def create_position_nodes
        nodes = []

        # Simple position
        nodes << Nokogiri::XML::Node.new("wp:simplePos", document).tap do |pos|
          pos["x"] = "0"
          pos["y"] = "0"
        end

        # Horizontal position
        nodes << Nokogiri::XML::Node.new("wp:positionH", document).tap do |posh|
          posh["relativeFrom"] = "column"
          posh << Nokogiri::XML::Node.new("wp:align", document).tap do |align|
            align.content = "center"
          end
        end

        # Vertical position
        nodes << Nokogiri::XML::Node.new("wp:positionV", document).tap do |posv|
          posv["relativeFrom"] = "paragraph"
          posv << Nokogiri::XML::Node.new("wp:posOffset", document).tap do |offset|
            offset.content = "0"
          end
        end

        Nokogiri::XML::NodeSet.new(document, nodes)
      end

      private def create_effect_extent_node
        Nokogiri::XML::Node.new("wp:effectExtent", document).tap do |effect_extent|
          effect_extent["l"] = "0"
          effect_extent["t"] = "0"
          effect_extent["r"] = "0"
          effect_extent["b"] = "0"
        end
      end

      private def create_doc_properties_node
        Nokogiri::XML::Node.new("wp:docPr", document).tap do |doc_pr|
          doc_pr["id"] = image_number
          doc_pr["name"] = image.name
          doc_pr["descr"] = image.name
        end
      end

      private def create_non_visual_properties_node
        Nokogiri::XML::Node.new("wp:cNvGraphicFramePr", document).tap do |c_nv_graphic_frame_pr|
          c_nv_graphic_frame_pr << Nokogiri::XML::Node.new("a:graphicFrameLocks", document).tap do |graphic_frame_locks|
            graphic_frame_locks["noChangeAspect"] = "1"
          end
        end
      end

      private def create_graphic_node(scaled_width, scaled_height)
        Nokogiri::XML::Node.new("a:graphic", document).tap do |graphic|
          graphic["xmlns:a"] = "http://schemas.openxmlformats.org/drawingml/2006/main"
          graphic << Nokogiri::XML::Node.new("a:graphicData", document).tap do |graphic_data|
            graphic_data["uri"] = "http://schemas.openxmlformats.org/drawingml/2006/picture"
            graphic_data << create_picture_node(scaled_width, scaled_height)
          end
        end
      end

      private def create_picture_node(scaled_width, scaled_height)
        Nokogiri::XML::Node.new("pic:pic", document).tap do |pic|
          pic["xmlns:pic"] = "http://schemas.openxmlformats.org/drawingml/2006/picture"

          # Non-visual picture properties
          pic << create_non_visual_picture_properties

          # Blip fill (the image reference)
          pic << create_blip_fill

          # Shape properties
          pic << create_shape_properties(scaled_width, scaled_height)
        end
      end

      private def create_non_visual_picture_properties
        Nokogiri::XML::Node.new("pic:nvPicPr", document).tap do |nv_pic_pr|
          nv_pic_pr << Nokogiri::XML::Node.new("pic:cNvPr", document).tap do |c_nv_pr|
            c_nv_pr["id"] = image_number
            c_nv_pr["name"] = image.name
            c_nv_pr["descr"] = image.name
          end
          nv_pic_pr << Nokogiri::XML::Node.new("pic:cNvPicPr", document).tap do |c_nv_pic_pr|
            c_nv_pic_pr << Nokogiri::XML::Node.new("a:picLocks", document).tap do |pic_locks|
              pic_locks["noChangeAspect"] = "1"
              pic_locks["noChangeArrowheads"] = "0"
            end
          end
        end
      end

      private def create_blip_fill
        Nokogiri::XML::Node.new("pic:blipFill", document).tap do |blip_fill|
          blip_fill << Nokogiri::XML::Node.new("a:blip", document).tap do |blip|
            blip["r:embed"] = node_id
            blip["cstate"] = "print"
          end
          blip_fill << Nokogiri::XML::Node.new("a:stretch", document).tap do |stretch|
            stretch << Nokogiri::XML::Node.new("a:fillRect", document)
          end
        end
      end

      private def create_shape_properties(scaled_width, scaled_height)
        Nokogiri::XML::Node.new("pic:spPr", document).tap do |sp_pr|
          sp_pr["bwMode"] = "auto"
          # Transform
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
          # Pre-set geometry
          sp_pr << Nokogiri::XML::Node.new("a:prstGeom", document).tap do |prst_geom|
            prst_geom["prst"] = "rect"
            prst_geom << Nokogiri::XML::Node.new("a:avLst", document)
          end
          # No fill
          sp_pr << Nokogiri::XML::Node.new("a:noFill", document)
        end
      end

      private def usable_dimensions_from(document)
        # Get page dimensions
        page_width = (document.at_xpath("//w:sectPr/w:pgSz/@w:w")&.value || Image::DEFAULT_PAGE_WIDTH).to_i
        page_height = (document.at_xpath("//w:sectPr/w:pgSz/@w:h")&.value || Image::DEFAULT_PAGE_HEIGHT).to_i

        # Convert to EMU
        width_emu = page_width * Image::TWENTIETHS_OF_A_POINT_TO_EMU
        height_emu = page_height * Image::TWENTIETHS_OF_A_POINT_TO_EMU

        # Account for margins
        left_margin = (document.at_xpath("//w:sectPr/w:pgMar/@w:left")&.value || 0).to_i * Image::TWENTIETHS_OF_A_POINT_TO_EMU
        right_margin = (document.at_xpath("//w:sectPr/w:pgMar/@w:right")&.value || 0).to_i * Image::TWENTIETHS_OF_A_POINT_TO_EMU
        top_margin = (document.at_xpath("//w:sectPr/w:pgMar/@w:top")&.value || 0).to_i * Image::TWENTIETHS_OF_A_POINT_TO_EMU
        bottom_margin = (document.at_xpath("//w:sectPr/w:pgMar/@w:bottom")&.value || 0).to_i * Image::TWENTIETHS_OF_A_POINT_TO_EMU

        # If no margins found, use defaults
        if left_margin == 0 && right_margin == 0
          left_margin = right_margin = Image::DEFAULT_MARGIN_IN_EMU
        end

        if top_margin == 0 && bottom_margin == 0
          top_margin = bottom_margin = Image::DEFAULT_MARGIN_IN_EMU
        end

        # Usable area
        usable_width = width_emu - left_margin - right_margin
        usable_height = height_emu - top_margin - bottom_margin

        # Add a small buffer (10%) to ensure image fits
        usable_width = (usable_width * 0.9).to_i
        usable_height = (usable_height * 0.9).to_i

        [usable_width, usable_height]
      end

      private def scale_dimensions(width_emu, height_emu, max_width_emu, max_height_emu)
        # Ensure we're working with EMU values throughout
        width_ratio = max_width_emu.to_f / width_emu
        height_ratio = max_height_emu.to_f / height_emu

        # Take the smaller ratio to ensure image fits within boundaries
        # but never scale up (keep at 1.0 if the image is smaller than available space)
        scale = [width_ratio, height_ratio, 1.0].min

        # Apply scaling factor and ensure integer values
        [(width_emu * scale).to_i, (height_emu * scale).to_i]
      end
    end
  end
end
