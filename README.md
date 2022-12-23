# convertworddoc
### A helpful tool to easily and safely convert word documents into a new format.
#### *Note: Primarily created for the purpose of converting outdated .doc documents into .docx and .txt for analysis using python-docx and regular expression parsing.*

### How To Use This:

I've included an example `main.py` script that uses the methods provided to find and covert a group of documents inside the given directory.

There are two main portions to converting documents: finding them, converting them. There is a method `src.convertdocs.method.create_doc_list()` built to find documents 
of a specific type provided for use, and two other functions to either convert a single document, or a list of documents. 
Please note that absolute paths are required.

After gathering the document file paths, you can use the `src.convertdocs.method.convert_documents()` or 
`src.convertdocs.method.convert_document()` to then convert all of those documents into a new file format.

Note: They will be saved in the same directory. Currently recursive searching and conversion isn't supported.

### Interesting Info:

This project was created originally to assist a request I had at work to convert over 1,000 Word Documents into a new 
format based on an updated template. Being used to Python and preferring to never touch VB or Macros, I immediately 
looked for a package that would help me do so, and found python-docx available for the job. Unfortunately, .doc files are 
completely incompatible because they are not OpenXML documents. After some quick research, I found this method and hope to 
create a useful package for someone else to quickly be able to download and move on with what they really want to do.



__References:__

pywin32 pypi:
https://pypi.org/project/pywin32/

pywin32 docs:
https://mhammond.github.io/pywin32/

WdSaveFormat Enumeration:
https://learn.microsoft.com/en-us/office/vba/api/Word.WdSaveFormat

OpenXML File Format:
https://learn.microsoft.com/en-us/office/open-xml/understanding-the-open-xml-file-formats