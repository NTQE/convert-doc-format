### UPDATE: 
Looks like I missed another potential solution in my hurry to complete this project. I've since found another solution that I haven't tested yet.

The Office Conversion Tool:

https://learn.microsoft.com/en-us/previous-versions/office/office-2010/cc179019(v=office.14)?redirectedfrom=MSDN

This method is likely faster, but I haven't test that. Still, it's interesting to know that something like that exists and why 
it wasn't pulling up higher in the search results on Google. Either way, this script was interesting to write and allows saving in j.txt and other formats as well.
I would have still needed this to convert all the documents to .txt for parsing due to some smart tags embedded in the document
that python-docx wasn't able to grab.

# convertworddoc
### A helpful tool to easily and safely convert word documents into a new format.
#### *Note: Primarily created for the purpose of converting outdated .doc documents into .docx and .txt for analysis using python-docx and regular expression parsing.*

### How To Use This:

I've included an example `main.py` script that uses the methods provided to find and covert a group of documents inside the given directory.

Note: This can't be done asynchronously as there are limitations in pywin32 and potentially win32 API's. Also, this method uses 
the SaveAs method on the active document, so only one document can be active at a time to ensure each document is handled properly. 
Expect an overhead of about 4-5 seconds for word, and somewhere around 1 second for each document, maybe less.

There are two main portions to converting documents: finding them, converting them. There is a method `src.convertdocs.method.create_doc_list()` built to find documents 
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