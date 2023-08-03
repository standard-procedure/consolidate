# Standard::Procedure::Consolidate

A simple gem for performing mailmerge on Microsoft Word .docx files.

Important: I can't claim the credit for this - I found [this gist](https://gist.github.com/ericmason/7200448) and have just adapted it.

It's pretty simple, so it probably won't work with complex Word documents, but it does what I need.  YMMV.


## Installation

    $ bundle add standard-procedure-consolidate

## Usage

To list the merge fields within a document:

```ruby
Consolidate::Docx::Merge.open "/path/to/docx" do |merge|
  puts merge.examine
end and nil
```
To perform a merge, replacing merge fields with supplied values:

```ruby
Consolidate::Docx::Merge.open "/path/to/docx" do |merge|
  merge.data "Name" => "Alice Aadvark", "Company" => "TinyCo", "Job_Title" => "CEO"
  merge.write_to "/path/to/output.docx"
end
```

NOTE: The merge fields are case-sensitive - which is why they should be supplied as strings (using the older `{ "key" => "value" }` style ruby hash).

## Development

The repo contains a .devcontainer folder - this contains instructions for a development container that has everything needed to build the project.  Once the container has started, you can use `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

`bundle exec rake install` will install the gem on your local machine (obviously not from within the devcontainer though). To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/standard-procedure/standard-procedure-consolidate.

When adding commit messages, please explain _why_ the change is being made.

When submitting a pull request, please ensure that there is an RSpec detailing how the feature works and please explain, in the pull request itself, the reasoning behind adding the feature.

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the Standard::Procedure::Consolidate project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/standard-procedure/standard-procedure-consolidate/blob/main/CODE_OF_CONDUCT.md).
