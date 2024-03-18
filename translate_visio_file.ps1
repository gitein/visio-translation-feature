# ⚪ NOTES FOR USERS

# Translate your diagrams in a Visio file using Office 365
# You must edit the file you'd translate to include the suffix"-tr" at the end of the visio file name, otherwise the script will exit and won't proceed (this is just not work on a wrong file)

# If criteria met, it creates a folder with "_copy" at the end which later will be exported as a copy .vsdx file tha holds the translation

# Once MS Word 365 opens proceed with the instructions to tranlate the text and update the TTT.txt file which will be created by the script itself and then gets deleted after exporting the final visio file.

# It's just about adding the text of TTT.txt to Word, translate it and then pasting back to the TTT.txt file, and the script does the heavy job on your behalf

# Thanks for using my solution


# _______________________________________
# 🔴 CODE STARTS HERE BELOW
# _______________________________________


# Get the directory path where the script is located
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path


# Find all files matching the criteria
$files = Get-ChildItem -Path $scriptDirectory -Filter "*-tr*.vsdx"


# Check if there are redundant files
if ($files.Count -gt 1) {
    Write-Host "There are redundant files matching the criteria:"
    $files | ForEach-Object { Write-Host $_.FullName }


    Write-Host "Please delete the redundant files manually and try again." -ForegroundColor Red
    Write-Host "Exiting script..."
    exit
} elseif ($files.Count -eq 0) {
    Write-Host "No files matching the criteria found in the directory." -ForegroundColor Red
    Write-Host "Exiting script..."
    exit
} else {
    Write-Host "File found: $($files[0].FullName)" -ForegroundColor Green
}


# PROCESSES VISIO FILE AND UNZIPs IT
# REQUIRES: A .VSDX FORMAT FILE IN PLACE
# _______________________________________

# Get the .vsdx file in the current directory

$vsdxFile = Get-ChildItem -Path $PSScriptRoot -Filter "*.vsdx" | Select-Object -First 1


# Check if a .vsdx file exists
if ($null -ne $vsdxFile) {
    # Create a copy of the .vsdx file
    $newFileName = "$($vsdxFile.BaseName)_copy.vsdx"
    Copy-Item -Path $vsdxFile.FullName -Destination $newFileName -Force


    # Rename the copied file to .zip format
    $zipFileName = ($newFileName -replace ".vsdx$", ".zip")
    Rename-Item -Path $newFileName -NewName $zipFileName
   
    Write-Host "File duplicated and renamed to .zip format successfully."  -ForegroundColor Green


    # Create a new folder
    $extractedFolder = New-Item -ItemType Directory -Path $vsdxFile.DirectoryName -Name "$($vsdxFile.BaseName)_copy"


    # Move the .zip file to the new folder
    Move-Item -Path $zipFileName -Destination $extractedFolder.FullName -Force


    Write-Host "Moved $zipFileName to $($extractedFolder.FullName)"


    # Extract the .zip file
    Expand-Archive -Path "$($extractedFolder.FullName)\$zipFileName" -DestinationPath $extractedFolder.FullName -Force
    Write-Host "Extracted $zipFileName to $($extractedFolder.FullName)"


    # Delete the .zip file inside the folder
    Remove-Item -Path "$($extractedFolder.FullName)\$zipFileName" -Force
    Write-Host "Deleted $zipFileName from $($extractedFolder.FullName)"  -ForegroundColor Green
} else {
    Write-Host "No .vsdx file found in the current directory." -ForegroundColor Red
}



# Creates A TXT FILE WITH ALL TEXT CONTENT IN CORRESPONDING .VSDX FILE
# REQUIRES: EXTRACTED VISIO FOLDER WITH >> PAGE1.XML << FILE IN IT
# _______________________________________

# Get all XML files recursively in the current directory
$xmlFiles = Get-ChildItem -Path $PSScriptRoot -Filter "page1.xml" -Recurse

foreach ($xmlFile in $xmlFiles) {
    # Load the XML file
    [xml]$xml = Get-Content $xmlFile.FullName

    # Define a recursive function to extract text content of Text elements
    function Extract-TextContent {
        param (
            [System.Xml.XmlNode]$node
        )

        $textContent = ""

        foreach ($childNode in $node.ChildNodes) {
            # If the node is a Text element, append its text content to the variable
            if ($childNode.LocalName -eq "Text") {
                $textContent += "$($childNode.InnerText)`r`n"  # Append text content with a new line
            }

            # Recursively call the function for child elements
            if ($childNode.HasChildNodes) {
                $textContent += Extract-TextContent -node $childNode
            }
        }

        return $textContent
    }

    # Call the function to extract text content of Text elements recursively
    $textContent = Extract-TextContent -node $xml

    # Construct the output file path in the same directory as the script
    $outputFilePath = Join-Path -Path $PSScriptRoot -ChildPath "TTT.txt"

    # Write the extracted text content to a file
    $textContent | Set-Content -Path $outputFilePath
    Write-Host "Created TTT.txt file successfully" -ForegroundColor Green

}

# OPENS WORD/EXCEL 365 TO PROCEED WITH TRANSALTION
# REQUIRES: INSTALLED OFFICE 365 VERSION
# _______________________________________


# Connect to Word application
$word = New-Object -ComObject Word.Application
$word.Visible = $true

# Create a new document
$word.Documents.Add()

Start-Sleep -Seconds 7  # Adjust the delay if necessary

#_________________________________

# Write text to the document
$word.Selection.TypeText("Add the text content of TTT.txt file and replace it here with this text, then select all text in the document (Ctrl + A) then hit Shift + Alt + F7 to open the translator in App and translate the text, then copy the translation back into the TTT.txt file, save the txt file, then proceed with the PowerShell script with Y for yes")


Start-Sleep -Seconds 1  # Adjust the delay if necessary

# Condition check before proceeding


# Function to prompt user and get input
function Prompt-YesNo {
    $input = Read-Host "Enter 'Y' for Yes or 'N' for No"
    if ($input -eq 'Y' -or $input -eq 'y') {
        return $true
    } elseif ($input -eq 'N' -or $input -eq 'n') {
        return $false
    } else {
        Write-Host "Invalid input. Please enter 'Y' for Yes or 'N' for No."
        Prompt-YesNo
    }
}

# Loop until user prompts Yes


# Function to prompt user and get input
function Prompt-Yes {
    $input = Read-Host "Enter 'Y' for Yes if TTT.TXT was updated with the translated text"
    if ($input -eq 'Y' -or $input -eq 'y') {
        return $true
    } else {
        return $false
    }
}

# Simple loop waiting for user to prompt 'yes'
while (-not (Prompt-Yes)) {
    Write-Host "Invalid input. Please enter 'Y' for Yes"  -ForegroundColor Red
}

# Proceed with the next commandlet here



#  End of loop

# Check if TTT.txt exists
$tttFilePath = Join-Path -Path $PSScriptRoot -ChildPath "TTT.txt"
if (Test-Path $tttFilePath) {
    Write-Host "TTT.txt file found. Processing translations..." -ForegroundColor Yellow
    
    # Read translated content from TTT.txt
    $translatedLines = Get-Content $tttFilePath | Where-Object { $_ -ne "" }

    # Load the XML file
    [xml]$xml = Get-Content $xmlFile.FullName

    # Define a function to update text content of Text elements
    function Update-TextContent {
        param (
            [System.Xml.XmlNode]$node,
            [ref]$translatedIndex
        )

        foreach ($childNode in $node.ChildNodes) {
            # If the node is a Text element, update its text content with translated content
            if ($childNode.LocalName -eq "Text") {
                if ($translatedIndex.Value -lt $translatedLines.Count) {
                    # Get the translated line corresponding to this text element
                    $translatedText = $translatedLines[$translatedIndex.Value]
                    $translatedIndex.Value++
                    
                    # Update text content with translated content
                    $childNode.InnerText = $translatedText.Trim()
                } else {
                    Write-Host "Not enough translations in TTT.txt to update all text elements. Aborting update." -ForegroundColor Yellow
                    return
                }
            }

            # Recursively call the function for child elements
            if ($childNode.HasChildNodes) {
                Update-TextContent -node $childNode -translatedIndex $translatedIndex
            }
        }
    }

    # Initialize translated index
    $translatedIndex = [ref]0

    # Call the function to update text content of Text elements recursively
    Update-TextContent -node $xml -translatedIndex $translatedIndex

    # Save the updated XML to the original file path
    $xml.Save($xmlFile.FullName)
    Write-Host "Updated page1.xml file with translated content successfully" -ForegroundColor Green
} else {
    Write-Host "TTT.txt file not found. Please run the script again after creating the TTT.txt file." -ForegroundColor Yellow
}


# EXPORTS FINAL VISIO FILE WITH TRANSLATED CONTENT
# REQUIRES: AN ONLY FOLDER IN PLACE TO COMPRESS AND THEN CONVERT TO .VSDX FILE
# 🔴 WARNING: CONSUMES THE ONLY FILE/FOLDER, I.E NO COPY CREATED
# _______________________________________

# Get the folder containing the extracted files
$extractedFolder = Get-ChildItem -Path $PSScriptRoot -Filter "*_copy" -Directory | Select-Object -First 1

# Check if the extracted folder exists
if ($extractedFolder -ne $null -and (Test-Path $extractedFolder.FullName)) {
    Write-Host "Extracted folder found: $($extractedFolder.FullName)"
    
    # Get the folder containing the .vsdx file
    $vsdxFolder = $extractedFolder.FullName
    Write-Host "Folder to be zipped: $($vsdxFolder)" -ForegroundColor Blue

    # Create a new .zip file
    $zipFileName = "$($extractedFolder.Name).zip"
    $zipFilePath = Join-Path -Path $PSScriptRoot -ChildPath $zipFileName

    # Compress the folder contents into a zip file
    Compress-Archive -Path $vsdxFolder\* -DestinationPath $zipFilePath -Force
    Write-Host "Created $zipFileName in $($PSScriptRoot)"

    # Remove the extracted folder
    Remove-Item -Path $extractedFolder.FullName -Recurse -Force
    Write-Host "Removed $($extractedFolder.FullName)"

    # Convert .zip to .txt
    $vsFileName = "$($extractedFolder.Name).vsdx"
    $vsFilePath = Join-Path -Path $PSScriptRoot -ChildPath $vsFileName

    # Read the contents of the .zip file as bytes and write to .vsdx file
    $zipBytes = [System.IO.File]::ReadAllBytes($zipFilePath)
    [System.IO.File]::WriteAllBytes($vsFilePath, $zipBytes)
    Write-Host "Converted $zipFileName to $vsFileName"

    # Remove the TTT.txt file
    Remove-Item -Path (Join-Path -Path $scriptDirectory -ChildPath "TTT.txt") -Force

    # Remove the zip file
    Remove-Item -Path $zipFilePath -Force
    Write-Host "Removed $zipFileName"

} else {
    Write-Host "Folder ending with '_copy' not found in $($PSScriptRoot)" -ForegroundColor Red
}
