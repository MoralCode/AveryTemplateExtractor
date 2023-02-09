# AveryTemplateExtractor

This little python script analyzes a `.docx` file containing a table (such as [avery's templates](https://www.avery.com/templates)) and prints out a series of key, value pairs representing the different parameters needed to recreate the template.

Written and tested on a linux machine.

## Setup and usage

1. `pipenv install` dependencies
2. `pipenv run python3 ./templateextractor.py ./path/to/docx/file.docx`

The file *must* be converted to docx first. this can be done with one of many commands, such as `unoconv -d document --format=docx *.doc`. Other GUI/CLI tools, like libreoffice/soffice, and pandoc can also be used for this.

