#~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*
#
# Decription:
# This script is used to prepare student submissions for the ImageCheck script.
# It first converts any Word document to PDF.
# Then it renames all documents to "Student Name.pdf"
# This script assumes that the documents were downloaded from D2L.
# You will probably have to set the $path variable below to match your
# file paths before you run this script.
#
# Developed by: Jeremy Dalby
#
#~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*

# **IMPORTANT** Modify this before running!
$path = "C:\Users\jerem\Downloads\Lab4"

# Convert any Word documents to PDF
$num_docs_converted = 0
$word_app = New-Object -ComObject Word.Application
$docs = Get-ChildItem -Path $path -Recurse -Filter *.doc? 

foreach ($doc in $docs)
{
    $document = $word_app.Documents.Open($doc.FullName)
    $pdf_filename = "$($doc.DirectoryName)\$($doc.BaseName).pdf"
    $document.SaveAs([ref] $pdf_filename, [ref] 17)
    $document.Close()
    Remove-Item $doc.FullName
    $num_docs_converted++
}
$word_app.Quit()

# Rename all PDF files as "Student Name.pdf"
$num_pdfs_renamed = 0
foreach ($item in Get-ChildItem $path)
{
    if ($item.Mode -match "^d")
    {
        $subs = Get-ChildItem $item.FullName
        $sub_num = 0
        foreach ($file in $subs)
        {
            # Make sure file is PDF
            if ($file.Name -match ".pdf$")
            {
                $split = $item.Name.Split("-")
                $student_name = $split[2].Trim()

                $new_name = "Oops.pdf"
                if ($subs.Length > 1)
                {
                    $new_name = $student_name + "_" + $sub_num + ".pdf"
                    $sub_num++
                }
                else
                {
                    $new_name = $student_name + ".pdf"
                }
            
                Rename-Item -Path $file.FullName -NewName $new_name
                $num_pdfs_renamed++
            }
            else
            {
                echo ("Non-PDF document found: " + $file.FullName)
            }
        }
    }
}

echo "Processing complete"
echo ("Number of Word documents converted to PDF = " + $num_docs_converted)
echo ("Number of PDF documents renamed = " + $num_pdfs_renamed)