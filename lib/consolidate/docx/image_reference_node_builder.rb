# frozen_string_literal: true

require "zip"
require "nokogiri"

module Consolidate
  module Docx
    class ImageReferenceNodeBuilder < Data.define(:field_name, :image, :node_id, :document)
      def call
        Nokogiri::XML::Node.new("w:drawing", document).tap do |drawing|
          drawing["xmlns:a"] = "http://schemas.openxmlformats.org/drawingml/2006/main"
          drawing << Nokogiri::XML::Node.new("wp:inline", document).tap do |inline|
            inline["distT"] = "0"
            inline["distB"] = "0"
            inline["distL"] = "0"
            inline["distR"] = "0"
            inline << Nokogiri::XML::Node.new("wp:extend", document).tap do |extent|
              extent["cx"] = image.width.to_s
              extent["cy"] = image.height.to_s
            end
            inline << Nokogiri::XML::Node.new("wp:effectExtent", document).tap do |effect_extent|
              effect_extent["l"] = "0"
              effect_extent["t"] = "0"
              effect_extent["r"] = "0"
              effect_extent["b"] = "0"
            end
            inline << Nokogiri::XML::Node.new("wp:docPr", document).tap do |doc_pr|
              doc_pr["id"] = Time.now.to_i.to_s
              doc_pr["name"] = "officeArt object"
              doc_pr["descr"] = image.name
            end
            inline << Nokogiri::XML::Node.new("wp:cNvGraphicFramePr", document)
            inline << Nokogiri::XML::Node.new("a:graphic", document).tap do |graphic|
              graphic["xmlns:a"] = "http://schemas.openxmlformats.org/drawingml/2006/main"
              graphic << Nokogiri::XML::Node.new("a:graphicData", document).tap do |graphic_data|
                graphic_data["uri"] = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                graphic_data << Nokogiri::XML::Node.new("pic:pic", document).tap do |pic|
                  pic["xmlns:pic"] = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                  pic << Nokogiri::XML::Node.new("pic:nvPicPr", document).tap do |nv_pic_pr|
                    nv_pic_pr << Nokogiri::XML::Node.new("pic:cNvPr", document).tap do |c_nv_pr|
                      c_nv_pr["id"] = Time.now.to_i.to_s
                      c_nv_pr["name"] = image.name
                      c_nv_pr["descr"] = image.name
                      c_nv_pr << Nokogiri::XML::Node.new("pic:nvPicPr", document).tap do |c_nv_pic_pr|
                        c_nv_pic_pr << Nokogiri::XML::Node.new("a:picLocks", document).tap do |pic_locks|
                          pic_locks["noChangeAspect"] = "1"
                        end
                      end
                    end
                  end
                  pic << Nokogiri::XML::Node.new("pic:blipFill", document).tap do |blip_fill|
                    blip_fill << Nokogiri::XML::Node.new("a:blip", document).tap do |blip|
                      blip["r:embed"] = node_id
                      blip << Nokogiri::XML::Node.new("a:extLst", document)
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
                        ext["cx"] = image.width.to_s
                        ext["cy"] = image.height.to_s
                      end
                    end
                    sp_pr << Nokogiri::XML::Node.new("a:prstGeom", document).tap do |prst_geom|
                      prst_geom["prst"] = "rect"
                      prst_geom << Nokogiri::XML::Node.new("a:avLst", document)
                    end
                    sp_pr << Nokogiri::XML::Node.new("a:ln", document).tap do |ln|
                      ln["w"] = 12700
                      ln["cap"] = "flat"
                      ln << Nokogiri::XML::Node.new("a:noFill", document)
                      ln << Nokogiri::XML::Node.new("a:miter", document).tap do |miter|
                        miter["lim"] = 400000
                      end
                    end
                    sp_pr << Nokogiri::XML::Node.new("a:effectLst", document)
                  end
                end
              end
            end
          end
        end
      end
    end
  end
end
