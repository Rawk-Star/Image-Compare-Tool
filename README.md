# Image-Compare-Tool

Description:
------------
These PowerShell scripts can be used to compare the images within multiple PDF or Word documents to determine which documents contain identical images.  It is useful for checking which students have shared screenshots for lab activities that were supposed to have been done independently.

The FileRename script will first convert any Word documents to PDF format, and then rename all PDF documents to "Student Name.pdf".  It assumes that the files have been downloaded from a D2L dropbox.

The ImageCheck script will produce a CSV file that lists all images submitted by each student cross-referenced to the students that have submitted the identical image.

Prerequisites:
--------------
- PowerShell
- Microsoft Word
- Any Internet browser
- Internet connection

Instructions:
-------------

1. Download all student submissions from D2L.
- Go to the D2L dropbox folder
- Go to the Files tab.
- Scroll to the bottom and make sure all files are listed. Use the dropdown in the bottom right corner to select "200 per page" if necessary.
- Select all files and click "Download". Save the ZIP file.
- Extract the zip file.

2. Run the FileRename script
- Open the FileRename.ps1 script in Windows PowerShell ISE.
- Modify the path in the $path variable at the top of the script to match the patch where you extracted the ZIP file in step 1.
- Run the script.  It might take a few mintues to run.  Execution time is largely dependent on how many files are to be converted from Word to PDF.
- When script execution finishes, confirm that the number of PDF files renamed seems reasonable.

3. Extract the images from the PDFs
- Go to https://tools.pdf24.org/en/extract-images
- Drag and drop the folder where you extracted the ZIP file in step 1 into the box on the webpage.  Wait for the PDF files to be uploaded.
- Click the "Extract Images" button at the bottom of the webpage box. Wait.
- Click the "Download" button. Save the ZIP file.
- Extract the ZIP file.

4. Run the ImageCheck script
- Open the ImageCheck.ps1 script in Windows PowerShell ISE.
- Modify the path in the $dir variable at the top of the script to match the patch where you extracted the ZIP file in step 3.
- Modify the path in the $CSV variable at the top of the script to match the location and file name where you would like the image check report to be saved.
- Run the script.  It should only take several seconds to run.  It should produce a CSV file.

5. Open the CSV file in Excel (or whatever spreadsheet program you prefer)
- Filter the data (Data tab > Filter)
- Using the filter on the "Image Also Used By" column, exclude (Blanks). Each row that is displayed represents one image that has been shared by 2 or more students.

6. Penalize students for plagiarism as you see fit!