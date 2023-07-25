# frozen_string_literal: true

require_relative "lib/consolidate/version"

Gem::Specification.new do |spec|
  spec.name = "standard-procedure-consolidate"
  spec.version = Consolidate::VERSION
  spec.authors = ["Rahoul Baruah"]
  spec.email = ["rahoulb@standardprocedure.app"]

  spec.summary = "Simple ruby mailmerge for Microsoft Word .docx files."
  spec.description = "Simple ruby mailmerge for Microsoft Word .docx files."
  spec.homepage = "https://github.com/standard-procedure/standard-procedure-consolidate"
  spec.license = "MIT"
  spec.required_ruby_version = ">= 2.6.0"

  spec.metadata["allowed_push_host"] = "https://rubygems.org"

  spec.metadata["homepage_uri"] = spec.homepage
  spec.metadata["source_code_uri"] = "https://github.com/standard-procedure/standard-procedure-consolidate"
  spec.metadata["changelog_uri"] = "https://github.com/standard-procedure/standard-procedure-consolidate/blob/main/CHANGELOG.md"

  spec.files = Dir.chdir(__dir__) do
    `git ls-files -z`.split("\x0").reject do |f|
      (File.expand_path(f) == __FILE__) ||
        f.start_with?(*%w[bin/ test/ spec/ features/ .git .circleci appveyor Gemfile])
    end
  end
  spec.bindir = "exe"
  spec.executables = spec.files.grep(%r{\Aexe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_dependency "rubyzip"
  spec.add_dependency "nokogiri"
end
