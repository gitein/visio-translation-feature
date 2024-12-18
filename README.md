# 🔷 Visio-Translation-Feature via PowerShell (Workaround)

### ⚪ Requirements:

- Installed and activated Office 365 version on a Windows machine.
- You first must edit the file you wish to translate to include the suffix "-tr" at the end of the Visio file name, e.g. `myfile-tr.vsdx`. Otherwise, the script will exit and won't proceed. (this is just to avoid working accidentally on the wrong file).
- If it meets the criteria, it will run smoothly and create a folder with suffix "_copy", which will later be exported as a `.vsdx` file with the translation.
- Use the file version `translate_visio_stable.ps1` if you don't wish to automatically open a new blank word document every single time you run the script, whereas `translate_visio_file.ps1`doesn't opens MS Word during runtime.

## Steps:

- Create an empty folder and move the powershell script file and the file you wish to translate to the same place together.
- Copy the path of the script with "Shift + right clicking" the powershell script file, and select "Copy as path".
- Run PowerShell as administrator.
-  Change current directory to the path where the script is saved `powershell.exe -File "PATH_TO_SCRIPT_WITH_FILE.vsdx"`, and replace the parameter with the real value of the path including the file name and extension, and don't remove the quotes.
  
- You may also need to temporarily override execution policy in your PowerShell environment before you run the script if you get an error that refers to blocking scripts of untrusted publishers. If so, then consider using the command `Unblock-File -Path "PATH_TO_SCRIPT_WITH_FILE.vsdx"` or `Set-ExecutionPolicy -Scope CurrentUser <Policy>` in case it failed to run using the previous method and then feel free to reset to everything to defaults after the job is done, and this is for your security, of course.
  
- It’s an important matter to bear in mind, so please ensure you reset the execution policy, if you want to view the current policy use the command `Get-ExecutionPolicy -List`, policy `RemoteSigned` is considered moderate and safe.
- Read the Microsoft documentation for more information [Link](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.4).

- It might throw some red error, ignore them, they just relate to error when connecting to COM server of Microsoft that is responsible for initiating a MS Word instance, it will still work fine as it's already tested in many environments.

- Translate your diagrams in the Visio file using Word 365, Copy and paste the content text out of the TTT.txt the script will create for you, paste into Word, select it, and hit "Shift + Alt + F7" to translate.
- Then paste the translated text of Word back into the TTT.txt file, confirm that TTT.txt file was updated by inputting "Y" in the powershell pipeline and the script will take care of the heavy job for you.

- That's pretty much it.
- Thanks for using my solution.
