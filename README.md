Automate ABBYY FineReader
=========================

ABBYY FineReader service wasn't available for pirate download and I had a bunch of 
books to convert. Since this is market-leading software I didn't want to use shitty
Linux tools like Tesseract.

This script automates ABBYY's UI to process books one by one, in VBScript (yuck!)

Setup
=====

1. Install Windows XP ISO in a VM.
2. Create input, output and tmp dirs inside ~/PDF-Maker
2. Install ABBYY Finereader 11 and crack
3. Open ABBYY and create this task:

  Step 1: Create new document

  Step 2: Open Image / PDF. Process images from this folder "~/PDF-Maker/input"

  Step 3: Analyze. Analyze the layout automatically

  Step 4: Read. Uncheck "open the verification dialog box"

  Step 5: Save document. Save it to ~/PDF-Maker/tmp as output.pdf


Running
=======

Run "cscript autopdf.vbs" to start processing. Copy PNGs to "~/PDF-Maker/input" and
make a new text document containing the PDF file name, then rename it to "start.txt"
to begin processing.

See feeder.vbs for a script to keep adding new dirs to process. 

ABBYY is a bit flaky so it may die from time to time, so manual intervention may be
necessary.

