# frozen_string_literal: true

module Consolidate
  class Image
    attr_reader :name, :width, :height, :aspect_ratio, :dpi

    def initialize name:, width:, height:, path: nil, url: nil, contents: nil
      @name = name
      @width = width
      @height = height
      @path = path
      @url = url
      @contents = contents
      @aspect_ratio = width.to_f / height.to_f
      #  TODO: Read this from the contents
      @dpi = {x: 72, y: 72}
    end

    def to_s = name

    def contents = @contents ||= contents_from_path || contents_from_url

    private def contents_from_path = @path.nil? ? nil : File.read(@path)

    private def contents_from_url = @url.nil? ? nil : URI.open(@url).read # standard:disable Security/Open
  end
end
