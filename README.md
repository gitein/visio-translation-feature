# visio-translation-feature

# ⚪ NOTES FOR USERS

- Requirements: Office 365 installed and activated

- Create a empty folder and place the powershell script file and the file you wisht to translate in the same place together,
- Copy the path of the script by Shift + right clicking the powershell script file, and select "Copy as path"
- then run PowerShell as administrator and
- run the scritp with this formula {. "<PATH_OF_SCRIPT>"} without curly braces, but don't remove the quotes
- It might throw some red error, ignore them, it will still work fine as tested many times in many environments

- Translate your diagrams in a Visio file using Office 365.
- You must edit the file you'd translate to include the suffix "-tr" at the end of the Visio file name. Otherwise, the script will exit and won't proceed. (This just to make sure you're not working on the wrong file.)
- If criteria met, it creates a folder with "_copy" at the end which later will be exported as a copy .vsdx file that holds the translation.
- Once MS Word 365 opens, proceed with the instructions to translate the text and update the TTT.txt file which will be created by the script itself and then gets deleted after exporting the final Visio file.
- It's just about adding the text of TTT.txt to Word, translate it, and then pasting back to the TTT.txt file, and the script does the heavy job on your behalf.

- Thanks for using my solution.
