# Doctool
A tool to manipulate .docx files, which enables to
* remove the ennoying "protection" that blocks some features (e.g. to restrict the formatting)
* list all the contributors of track changes
* rename those authors (e.g. for correcting a mistake or achieve consistent and homogeneous naming policy. Thanks to [Alexandre Marquet](https://github.com/alexmrqt) for the original tips)
* make the document smaller by converting embedded EMF and PNG pictures to SVG or JPG (lossy ! Please check the output document !)

# Usage
This is for the moment a command-line tool. It should work on every computer with Python (this was tested on Linux (Manjaro), and on Windows with the [Anaconda](https://www.anaconda.com/products/individual) Python distribution)

* `python doctool.py example.docx removeprotection` : remove the "protection" of the document (e.g. in order to add new formatting or styles).
* `python doctool.py example.docx list_authors` : returns the list of authors in the track changes
* `python doctool.py example.docx change_authors "old1" "new1" "old2" "new2"...` : change authors names (from old to new. N.B. Don't forget to have quotation marks if a name include spaces). You may append `-o output_file.docx` in order to preserve the original file. Beware: changing several author names to a common one is not properly tested (behavior is unknown e.g. when several track changes from several authors overlap/conflict those authors are all renamed to a common one...)
* `python doctool.py example.docx slimfast [-o output.docx]` : Reduces the size of the .docx file by converting embedded images : PNG over 30kB are converted to JPG, and EMF are converted first to SVG (using [libemf2svg](https://github.com/kakwa/libemf2svg), which seems to produce good quality results. Only compiled and tested on Linux at the moment). Resulting SVG may be already significantly more lightweight than the original EMF in some case, and if it is still above 600kB the script will rasterize the SVG to JPG. Of course all of this is a lossy compression => use it at your own risk and check the result !
