# convertworddoc
### A helpful tool to easily and safely convert word documents into a new format.
#### *Note: Primarily created for the purpose of converting outdated .doc documents into .docx and .txt for analysis using python-docx and regular expression parsing.*

### How To Use This:

I've included an example `main.py` script that uses the methods provided to find and convert a group of documents inside the given directory.

Note: This can't be done asynchronously as there are limitations in pywin32 and win32 API's. Also, this method uses 
the SaveAs method on the active Word document, so only one document can be active at a time to ensure each document is handled properly. 
Expect an overhead of about 4-5 seconds for word, and somewhere around 1 second for each document, maybe less.

There are two main portions to converting documents: 1. finding them, 2. converting them. There is a method `src.convertdocs.method.create_doc_list()` built to find documents 
of a specific type provided for use, and two other functions to either convert a single document, or a list of documents. 
Please note that absolute paths are required.

After gathering the document file paths, you can use the `src.convertdocs.method.convert_documents()` or 
`src.convertdocs.method.convert_document()` to then convert all of those documents into a new file format.

Note: They will be saved in the same directory. Currently recursive searching and conversion isn't supported.


__References:__

pywin32 pypi:
https://pypi.org/project/pywin32/

pywin32 docs:
https://mhammond.github.io/pywin32/

WdSaveFormat Enumeration:
https://learn.microsoft.com/en-us/office/vba/api/Word.WdSaveFormat

OpenXML File Format:
https://learn.microsoft.com/en-us/office/open-xml/understanding-the-open-xml-file-formats


## UPDATE: 
Found another potential solution that I haven't tested yet. May or may not have suited the requirements of this project, but it doesn't seem to convert to .txt and that was required to parse the smart tags embedded in the documents. Not sure if this tool properly converts those tags.

The Office Conversion Tool:

https://learn.microsoft.com/en-us/previous-versions/office/office-2010/cc179019(v=office.14)?redirectedfrom=MSDN
