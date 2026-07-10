Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = folder
sh.Run """D:\Python313\pythonw.exe""" & " """ & folder & "\_wps_pdf_capture_test.py""", 1, False
