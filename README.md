# Templatr

Templatr is designed to take a Word file with placeholder text and replace it with values in a JSON Object. It supports replacement of text with paragraphs, tables, images, and “lists” of content. 

## Getting Started

Template requires two main libraries firstly it utilizes docx4j for word file manipulation and secondly it uses json-simple to read and parse the JSON file. When using this project you will need both of these libraries along with their dependences. I’d recommend you use apache maven to manage dependencies. You can use the pom.xml for the dependencies used in my project. 

## Usage

Templatr is incredibly simple to use, simply create your JSON file using the format shown in input.json, and pass in both the JSON file along with your word file to the constructor. Once you do this it will begin to replace the required information. You can then save the file by calling saveDocument() and passing in the path of the file you want to save. I’d recommend using a different name so it doesn’t overwrite your template file at least until you are satisfied with how it’s working.



### Notes on placeholders

Placeholders can have any format you like but I’d recommend something very distinct that normal uses won’t enter in a document. (I used the format of [:PLACEHOLDER_TEXT:]). 


