# Doctool
A tool to manipulate .docx files (thanks to [Alexandre Marquet](https://github.com/alexmrqt) and Vincent Durepaire for the original idea :)

This tool enables to
* remove the ennoying "protection" that blocks some features (e.g. to restrict the formatting)
* list all the contributors of track changes
* rename those authors (e.g. for correcting a mistake or achieve consistent and homogeneous naming policy)

# Usage
This is for the moment a command-line tool. It should work on every computer with Python (this was tested on Linux (Manjaro), and on Windows with the [Anaconda](https://www.anaconda.com/products/individual) Python distribution)

* `python doctool.py example.docx removeprotection` : remove the "protection" of the document (e.g. in order to add new formatting or styles).
* `python doctool.py example.docx list_authors` : returns the list of authors in the track changes
* `python doctool.py example.docx change_authors "old1" "new1" "old2" "new2"...` : change authors names (from old to new. N.B. Don't forget to have quotation marks if a name include spaces). You may append `-o output_file.docx` in order to preserve the original file
