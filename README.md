# 🔷 Visio-Translation-Feature via PowerShell (Workaround)

### ⚪ NOTES FOR USERS

- Requirements: Office 365 installed and activated. 
- You must edit the file you'd translate to include the suffix "-tr" at the end of the Visio file name. Otherwise, the script will exit and won't proceed. (This just to make sure you're not working on the wrong file.)
- If it meets the criteria, it creates a folder with "_copy" at the end which later will be exported as a .vsdx file that which encompasses the translation.
- Use `translate_visio_stable.ps1` which doesn't open MS Word automatically for more stable performance, whereas `translate_visio_file.ps1` opens MS Word during runtime.

## Steps:

- Create an empty folder and move the powershell script file and the file you wish to translate to the same place together.
- Copy the path of the script with "Shift + right clicking" the powershell script file, and select "Copy as path".
- Run PowerShell as administrator.
- You may need to override execution policy in PowerShell before you run the script in case you get an error that refers to scripts block due to untrusted resources. If so, consider using the command `Set-ExecutionPolicy -Scope CurrentUser <Policy>` and then feel free to reset to defaults after it does the job.
- Read the Microsoft documentation for more information [Link](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.4)

  
- Run the following command in PowerShell `. "PATH_TO_SCRIPT"` and replace the placeholder with the real path, but don't remove the quotes.
- It might throw some red error, ignore them, they just relate to error when connecting to COM server of Microsoft that is responsible for opening MS Word, it will still work fine as tested many times in many environments.

- Translate your diagrams in the Visio file using Word 365, Copy and paste the content text out of the TTT.txt the script will create for you, paste into Word, select it, and hit "Shift + Alt + F7" to translate.
- Then paste the translated text of Word back into the TTT.txt file, confirm that TTT.txt file was updated by inputting "Y" in the powershell pipeline and the script will take care of the heavy job for you.

- That's pretty much it.
- Thanks for using my solution.
