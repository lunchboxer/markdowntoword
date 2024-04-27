# Markdown to Word

A simple go program which copies strings from a markdown file to a word file using a template with placeholders. Placeholders are delimited using `{key}`. On the markdown side, the program looks for third level headings and definition lists to build the replacement map.

Labels for placeholders are kebab case and prefixed by the text of the previous second-level heading.
