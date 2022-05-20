# Excel-2-Doc

This program uses content from an excel spreadsheet (format.xls) and creates a Word Document (new_test.docx). The code uses the document excel2doc_template.docx as a template.

Python docx aswell as docxtpl is used to run the program.

Files:
- format.xls is the Excel Spreadsheet populated with Dummy Data.
    - This Dummy Data showcases the use of Headers, Sub-Headers, Paragraphs, and Tables.
- excel2doc_template.docx is the Word Document Template that the code uses to create the new Word Document.
    - multiple prefixes can be seen in this document such as {{document_title}}, and {{revision}}.
- new_test.docx is the newly created Word Document populated with the contents inside the Excel Spreadsheet.

This program automates the creation of documentation for a business setting. One constant issue I have noticed during my employment that kept arising was time spent working on documentation with a particular struggle with the layout of a word document. Issues such as Word freezing on employees while working, not being able to get the borders to fit correctly, and so on. Silly, tedious issues.

An additional benefit that this process allows is to categorize the data. For example, Procedural documentation can be categorized by RACI roles (Responsible, Accountable, Informed, Consulted) that help filter through the data so certain individuals can highlight their duties.

Another example, Function Design Documents can be inputted into a EXCEL file. When documentation needs to be signed off, users can simply run the program to create all Function Design Documents for each function he/she may have created during a Sprint. The business now has two versions of the data, one in xlsx format and another as electronically signed indivdual documents. The xlsx document can later be used to filter through to find out much needed information. This is a much more time efficient solution compared to reviewing Word Documents or PDFs one-by-one. This also allows developers to focus more on development and little on the layout of a document. 
