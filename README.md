[![GitHub release (latest by date)](https://img.shields.io/github/v/release/julienbordet/PayRaise)](https://github.com/julienbordet/MenuPing/releases)
![GitHub commit activity](https://img.shields.io/github/commit-activity/m/julienbordet/PayRaise)

# DocxTemplate
Simple Python script to generate word documents from an Excel file.

# Help

```
$ python app/docxtemplate.py -h
usage: docxtemplate.py [-h] [-prefix PREFIX] [-v] directory xlsx_data template

positional arguments:
  directory       target directory for generated files
  xlsx_data       excel data file
  template        docx template

optional arguments:
  -h, --help      show this help message and exit
  -prefix PREFIX  prefix for generated files
  -v              verbose mode (-vv for more)
```

# Exemple usage

```
$ python app/docxtemplate.py target/2022 target/2022/raises.xlsx template/letter.docx
```

* `target/2022` is the directory in which the docx files will be created
* `target/2022/raises.xlsx` is the excel file containing the core data
* `template/letter.docx` is the model docx file

## Docx template format

Use '***variable_name***' to insert into the template docx the value of the *variable_name* column from the Excel file.

## Edit variable in Excel file

The '.xlsx' file should contains :
* a header with the name of the variable that should be matched in the docx template
* the different values

**Note** : each first column value should be unique.

# Installation

Simply launch

```
$ chmod u+x install.sh
$ ./install.sh
```