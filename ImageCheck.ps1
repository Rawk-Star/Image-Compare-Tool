#~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*
#
# Decription:
# This script is used to determine who shared screenshots.
# It will produce a CSV file that lists all images submitted by each 
# student cross-referenced to the students that have submitted the 
# identical image.
# It assumes that the FileRename script has already been run.
# You will probably have to set the $dir and $CSV variables below 
# to match your file paths before you run this script.
#
# Developed by: Jeremy Dalby
#
#~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*~*

# **IMPORTANT** Modify these before running!
$dir = "C:\Users\jerem\Downloads\Lab4Images"
$CSV = "C:\Users\jerem\Downloads\Lab4ImageHashes.csv"

# Get hashes of all images
$all_imgs = Get-ChildItem -Recurse $dir | Get-FileHash

# Output header row to CSV file
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$headers = "Image Hash,Image File,Student Name,Image Also Used By"
$headers > $CSV

foreach ($base_img in $all_imgs)
{
    # Get the info for this image
    $result = $base_img.Hash.ToString()
    $result = $result + "," + $base_img.Path.ToString()
    $result = $result + "," + $base_img.Path.Split("\")[-2]
    
    #Determine which other students used the same image
    $img_also_used_by = ""
    foreach ($comp_img in $all_imgs)
    {
        if ($base_img -ne $comp_img)
        {
            if ($base_img.Hash -eq $comp_img.Hash)
            {
                $student = $comp_img.Path.Split("\")[-2]

                if ($img_also_used_by.Length -gt 0)
                {
                    $img_also_used_by = $img_also_used_by + " & " + $student
                }
                else
                {
                    $img_also_used_by = $student
                }
            }
        }
    }
    
    $result = $result + "," + $img_also_used_by
    $result >> $CSV
}

echo "Processing complete"