# Standard::Procedure::Consolidate

A simple gem for performing search and replace on Microsoft Word .docx files.

Important: I can't claim the credit for this - I found [this gist](https://gist.github.com/ericmason/7200448) and have adapted it for my needs.

## Search/Replace for field placeholders

If you have a Word .docx file that looks like this (ignoring formatting): 

```
Dear {{ first_name }},

Thank you for your purchase of {{ product }} on the {{ date }}.  We hope that with the proper care and attention it will give you years of happy use.  

Regards

{{ user_name }}
```

We have marked out the "fields" by using squiggly brackets - in a manner similar to (but simpler than) Liquid or Mustache templates.  In this example, we have fields for first_name, product, date and user_name.  

Consolidate reads your .docx file, locates these fields and then replaces them with the values you have supplied, writing the output to a new file.  

NOTE: These are not traditional Word "mail-merge fields" - these are just fragments of text that are within the Word document.  See the history section for why this does not work with merge-fields.  

## Usage

There is a command-line tool or a ruby library that you can include into your code.  

Using the ruby library: 

```ruby
Consolidate::Docx::Merge.open "/path/to/file.docx" do |doc|
  puts doc.field_names 

  doc.data first_name: "Alice", product: "Palm Pilot", date: "23rd January 2002", user_name: "Bob"
  doc.write_to "/path/to/merge-file.docx"
end 
```

Using the command line: 

```sh
examine /path/to/file.docx 

consolidate /path/to/file.docx /path/to/merge-file.docx first_name=Alice "product=Palm Pilot" "date=23rd January 2022" "user_name=Bob" 
```

If you want to see what the routine is doing you can add the `verbose` option.  

```sh
examine /path/to/file.docx verbose

consolidate /path/to/file.docx /path/to/merge-file.docx first_name=Alice "product=Palm Pilot" "date=23rd January 2022" "user_name=Bob" verbose
```

### History

Originally, this gem was intended to open a Word .docx file, find the mailmerge fields within it and then substitute new values.  

I managed to get a basic version of this working with a variety of different Word files and all seemed good.  Until my client reported that when they went to print the document, the mailmerge fields reappeared!  My best guess is that Word is thinking that print-time is a trigger for merging in a data source (for example, printing out a form letter to 200 customers), so all the substitution work that this gem does is then discarded and Word asks for the merge data again.  The frustrating thing is I can't figure out how Word keeps the references to the merge fields after they've been substituted.  

So instead, this does a simple search and replace - looking for fields within squiggly brackets and substituting them.  

### How it works

This is a bit sketchy and pieced together from the [code I found]((https://gist.github.com/ericmason/7200448)) and various bits of skimming through the published file format.

A .docx file is actually a .zip file that contains a number of .xml documents.  Some of these are for storing formatting information (fonts, styles and various metadata) and some are for storing the actual document contents.  

Consolidate looks for word/document.xml files, plus any files that match word/header*.xml, word/footer*.xml, word/footnotes*.xml and word/endnotes*.xml.  It parses the XML, looking for text nodes that contain squiggly brackets.  If it finds them, it then checks to see if we have supplied a data value for the matching field-name and replaces the contents of the node.  

## Installation

    $ bundle add standard-procedure-consolidate


## Development

The repo contains a .devcontainer folder - this contains instructions for a development container that has everything needed to build the project.  Once the container has started, you can use `bin/setup` to install dependencies. Then, run `bundle exec rake` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

`bundle exec rake install` will install the gem on your local machine (obviously not from within the devcontainer though). To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/standard-procedure/standard-procedure-consolidate.

When adding commit messages, please explain _why_ the change is being made.

When submitting a pull request, please ensure that there is an RSpec detailing how the feature works and please explain, in the pull request itself, the reasoning behind adding the feature.

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the Standard::Procedure::Consolidate project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/standard-procedure/standard-procedure-consolidate/blob/main/CODE_OF_CONDUCT.md).
