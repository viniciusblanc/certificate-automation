# certificate-automation
Documentation of my procedure to generate and delivery certificates

- The 1st step is compile the data to be used.
- The 2nd step is to create a direct mail.
- - If the direct mail is in MSWord, you can use a litle VBA script to create multiple PDFs individually named:

See the file **PDF_Generator.vb**

- - If you already have a single PDF file, you can use a PDF-Spliter like mine[^1]:

[^1]: For more details, see: https://github.com/viniciusblanc/pdf-splitter

See the file **split-pdf.py**

- The 3rd step is to upload all these files to a Google Drive.
- Once there, you must create a Google Sheet with the information that will be used to automate the emails.
- - On the Google Sheet, go to Extensions/App Script to put the following code:

See the file **sendEmails.js**
