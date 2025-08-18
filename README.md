Create your own categories and have the files in your subdirectory ai scanned and sorted/moved into category-titled subdirectories.  

Works on windows 8.1 machines, but with no GPU.

Instructions: download py file and run using below.
cmd> python AI_local_file_organizer.py

Top to bottom Gemini Canvas creation.  Modificaions in gemini canvas should work.

-----------------------------------------------------------------------------------

### Offline AI-Powered Python based File Sorter and Summarizer (with Windows 8 support)

This Python script uses AI to automatically read, summarize, and categorize your documents. It helps you quickly understand the content of your files and organize them into topic-based folders without manual effort.

The tool processes various file types, including PDF, DOCX, TXT, XLSX, and PPTX, using a local, offline AI model to ensure your data remains private.


**Key Features**

Automatic Summarization: Generates concise, bullet-point summaries of your documents using a t5-small model.

AI-Powered Categorization: Uses a zero-shot classification model (xtremedistil-l6-h256) to sort files into custom categories like "Health," "History," "Finance," etc.


##Flexible File Organization: ##

After categorization, you can automatically:

Move files to category folders.

Copy them.

Create Shortcuts.

Or do nothing and just view the analysis.


##Multiple File Sources:##

Process files by scanning an entire folder (and its subdirectories) or by providing a specific list of file paths in a text file.

Customizable: Easily edit categories, summary length, and file processing rules through a command-line interface.

Settings Memory: All your configuration choices are saved to a settings.json file, making subsequent runs fast and easy.

Offline and Private: Runs entirely on your local machine. No data is sent to the cloud.

------------------------------------------------------------

From console:

Do you want to [M]ove, [C]opy, create [S]hortcut, or [D]o nothing? [default: none]: s

--- Configuration Settings ---

Enter chunk length (words per chunk, default: 250):

Enter max summary length (words, default: 14): 17

Enter min summary length (words, default: 9):


--- PDF Settings (in 250-word chunks) ---

Enter number of initial chunks to skip (padding, default: 1): 0

Enter number of chunks to process after padding (default: 2): 2


--- Settings for Other File Types (in 250-word chunks) ---

Enter number of initial chunks to skip (padding, default: 0):

Enter number of chunks to process after padding (default: 2):

Enter confidence threshold for categorization (0-100, default: 11):

Tag files with category name in file properties? (y/n) [default: y]):

Enter confidence threshold for tagging (0-100, default: 16):

Add a prefix to organized files? (y/n) [default: n]):

Settings saved to settings.json for the next session.

Files will be organized without a prefix.

Step 11: Folder path is valid. Starting the batch process...

Step 4: Starting main processing for 5891 files...

Step 5: Processing file 1/5891 - bsdtar.1.pdf

Step 6: Extracting from chunk(s) 1 to 2...

Step 7: Text extracted successfully. Starting summarization...

Step 7.1.1: Summarizing chunk 1 of 2...

Step 7.1.2: Summarization for chunk 1 complete.

Step 7.1.1: Summarizing chunk 2 of 2...

Step 7.1.2: Summarization for chunk 2 complete.

Step 8: Categorizing summary offline...

Step 8.1: Sending summary to classifier model...

Step 8.2: Categorization complete.

Step 9: Final summary complete.

Error tagging file bsdtar.1.pdf: code=2: cannot open file 'C:\ProgramData\miniconda3\pkgs\libarchive-3.7.4-h9243413_0\Library\share\man\pdf\bsdtar.1.pdf': Permission denied

File: C:\ProgramData\miniconda3\pkgs\libarchive-3.7.4-h9243413_0\Library\share\man\pdf\bsdtar.1.pdf

Summary:

Category: database, data model, catalog, software, record, file (17.71%)

- tar creates and manipulates streaming archive files

- this implementation can extract

- tar c f - newfile @original.tar

- bsdtar.1
