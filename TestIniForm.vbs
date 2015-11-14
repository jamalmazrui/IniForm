Dim oIniForm
Dim sPath

sPath = "C:\IniForm\"
Set oIniForm = CreateObject("IniForm")
' msgbox IsObject(oIniForm)
' MsgBox TypeName(oIniForm)
i = oIniForm.RunForm(sPath)
' MsgBox i
MsgBox oIniForm.GetResult("Last Name"), 0, "GetResult Last Name"
i = oIniForm.ShowResults()
