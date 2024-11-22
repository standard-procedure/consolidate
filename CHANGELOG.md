## [0.3.1] - 2024-11-22

Ensure that the substituted nodes are reinserted correctly into the output document, attempting to restore formatting at the paragraph level (although it does lose formatting at lower levels than this - so-called "run" nodes which represent arbitrary spans of characters within the paragraph).


## [0.3.0] - 2024-11-21

Updated the code that examines the docx file for merge fields to deal with Word formatting tags being inserted in the middle of the merge fields.  

## [0.2.0] - 2023-09-13

Thrown away the mail-merge implementation and replaced it with a simple search/replace.  
Added in the command line utilities for examining and consolidating documents.  

## [0.1.4] - 2023-09-11

Updated which files get exclusions after crashes in production with some client files

## [0.1.3] - 2023-09-07

Customer had an issue where some fields were not merging correctly

- Altered the code so it tries to perform the substitution in all .xml files that aren't obviously configuration data
- Added verbose option to show what the code was doing
- Added a list of files to `examine`
- Still no tests for the header/footer support

## [0.1.2] - 2023-08-23

- Added untested support for headers and footers
- this work was done in a rush to deal with a customer requirement - tests to come soon

## [0.1.1] - 2023-08-02

- Added `examine` to list the merge fields within a document

## [0.1.0] - 2023-07-25

- Initial release

## [Unreleased]
