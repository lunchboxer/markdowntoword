# Markdown to Word

A simple go program which copies strings from a markdown file to a word file using a template with placeholders. Placeholders are delimited using `{key}`. On the markdown side, the program looks for third level headings and definition lists to build the replacement map.

Labels for placeholders are kebab case and prefixed by the text of the previous second-level heading.

## Set up

The program requires the following packages:

`github.com/lukasjarosch/go-docx`

To build and install the program:

`go install`

## Usage

Run the program with the markdown file as the first argument and the template word file as the second argument.
