# docxGenerator4Data

CLI utility co create a collection of DOCX-files by template and table of key-words

# Using

1. Download latest binary package of utility from [Releases](https://github.com/GeorgGrebenyuk/docxGenerator4Data/releases/latest);
2. Unzipped it;
3. Call `cmd.exe` from it's directory or navigate here with `cd /d` command;
4. (For testing) Run `net8.0\Sample\test.bat` to check, if application is worked correctly on your PC;

If need, install the system dependency (.NET 8.0 Runtime): https://dotnet.microsoft.com/en-us/download/dotnet/thank-you/runtime-8.0.21-windows-x64-installer.

## Create a DOCX-template file

Create a DOCX file, that have a specific words (aka anchors) in all places, which will be replacing with specific text. For simple documents you can use simple anchors -- WORD_1, WORD_2 and etc. For large documents recommend use GUID's (UUID too). 

## Create an anchor's file config

Create a text file. The first line must contains all anchor's names. Each of other strings will represent one DOCX file (under specific column with achor there is a replacing text). 

## Create a commandline arguments for app

* `DocxTemplatePath`: Absolute file path to *.DOCX template (with text anchors);
* `OutputDirectory` : Absolute file path to output directory (where result docx files will be created);
* `AnchorFilePath`: Absolute file path to table-file (one row represent one file, the column-values are the replaceable text for anchor's name in header-column);
* `AnchorFileSeparator`: A string -- separator of columns in input table-file (by AnchorFilePath);
* `NewFileName`: A name of createable DOCX file (default a file1, file2 ...). Use an anchor's name from AnchorFile to specify a need name from column; if file will exists with that name, it will be re-writed;

## Example of config and files

The content of docx's template file:

```txt
Person's name in WORD_2. It living in WORD_3.
```

The content of AnchorFile:

```txt
WORD_1	WORD_2	WORD_3
Doc 1	John Smitt	Berlin
Doc 23	Kristofer Nollan	New York
```

The command for utility:
```txt
docxGenerator4DataCLI DocxTemplatePath "E:\\Downloads\\TemplateFile.docx" OutputDirectory "E:\\Downloads\\Output" AnchorFilePath "E:\\Downloads\\AnchorFile.csv" AnchorFileSeparator "	" NewFileName = "WORD_1"
```

Sample files and script you can find at `net8.0\Sample`.