' IniForm
'Version 2.0
' November 12, 2015
' Copyright 2006 - 2015 by Jamal Mazrui
' GNU Lesser General Public License (LGPL)

#COMPILE DLL
#COM TLIB ON
#RESOURCE TYPELIB, 1, "IniForm.tlb"
#DIM ALL
' #DEBUG DISPLAY On
' #MESSAGES NOTIFY
#MESSAGES Command
' #OPTION ANSIAPI
#TOOLS Off
' #UNIQUE Var On

#COM NAME "IniForm", 1.4
#COM DOC "Build a Windows dialog and retrieve data from it via ini files"
#COM GUID GUID$("{586D59B9-B599-42C8-AE21-23B6A838BB22}")

#INCLUDE ONCE "Win32API.inc"
' %USEMACROS = 1
#Include Once "RichEdit.inc"
#Include Once "ShlWAPI.inc"
#Include Once "CommCtrl.inc"
#Include Once "comDlg32.inc"
#Include Once "WinINet.inc"
' #Include Once "mylib.inc"
#Include Once "EOIni.inc"

%ID_OK =1
%ID_CANCEL =2
%ID_STATUSBAR =999
%ID_F1 =%VK_F1 +%FVIRTKEY +1000
%ID_F3 =%VK_F3 +%FVIRTKEY +1000
%ID_SHIFT_F3 =%FSHIFT +%VK_F3 +%FVIRTKEY +1000
%ID_F5 =%VK_F5 +%FVIRTKEY +1000
%ID_CONTROL_A =%VK_A +%FCONTROL +%FVIRTKEY +1000
%ID_CONTROL_SHIFT_A =%VK_A +%FCONTROL +%FSHIFT +%FVIRTKEY +1000
%ID_CONTROL_F =%VK_F +%FCONTROL +%FVIRTKEY +1000
%ID_CONTROL_SHIFT_F =%VK_F +%FCONTROL +%FSHIFT +%FVIRTKEY +1000
%ID_CONTROL_R =%VK_R +%FCONTROL +%FVIRTKEY +1000
%ID_CONTROL_S =%VK_S +%FCONTROL +%FVIRTKEY +1000
%ID_CONTROL_ENTER =%VK_RETURN +%FCONTROL +%FVIRTKEY +1000

$CIniFormGuid = GUID$("{83A881DE-F4C4-41E8-A8DF-552C3FCCF4E1}")
$IIniFormGuid = GUID$("{1D60D806-82D0-40C6-B920-4E0BB18B8516}")

Global oIniForm as iDispatch
Global bDebugMode As Long

Type ThreadInfo
Source AS ASCIIZ * %MAX_PATH
Target AS ASCIIZ * %MAX_PATH
Parent As DWord
End Type

Type DialogBand
Height As Long
Top As Long
Bottom As Long
XSpace As Long
ButtonWidth As Long
CheckWidth As Long
RadioWidth As Long
CtlWidth As Long
DlgWidth As Long
Count As Long
End Type

Global a_dlgBand() As DialogBand, ti As ThreadInfo
Global a_band() As Long, a_name() As Asciiz * 81, a_id() As Asciiz * 11, a_control() As Asciiz * 11, a_type() As Asciiz * 11, a_mask() As Asciiz * 21
Global a_caption() As Asciiz * 81, a_range() As String, a_value() As String, a_select() As String, a_focus() As String, a_result() As String
Global a_align() As Asciiz * 11, a_left() As Asciiz * 11, a_top() As Asciiz * 11, a_width() As Asciiz * 11, a_height() As Asciiz * 11
Global a_fore() As Asciiz * 11, a_back() As Asciiz * 11, a_style() As Asciiz * 21, a_extend() As Asciiz * 11, a_tip() As Asciiz * 81, a_help() As String, a_misc() As Asciiz * 21
Global a() As String, a_multi() As Long, i As Long, i_id As Long, i_count As Long, c As Asciiz *81, s_item As String, s_list As String
Global a_keys() As ACCELAPI, h_dlg As Dword, i_result As Long, i_dlgLeft As Long, i_dlgTop As Long, i_dlgWidth As Long, i_dlgHeight As Long
Global i_dlgStyle As Long, i_dlgExtend As Long, s_dlgFont As Asciiz * 21, i_dlgPoint As Long, s_dlgMisc As Asciiz * 21, s_dlgInput As Asciiz * 11, s_dlgOutput As Asciiz * 11
Global x As Long, y As Long, xx As Long, yy As Long, xSpace As Long, YSpace As Long, i_labelPad As Long, i_borderPad As Long
Global i_labelWidth As Long, i_labelHeight As Long, i_buttonWidth As Long, i_buttonHeight As Long
Global i_CheckWidth As Long, i_checkHeight As Long, i_radioWidth As Long, i_radioHeight As Long
Global i_listWidth As Long, i_listHeight As Long, i_multiWidth As Long, i_multiHeight As Long
Global i_editWidth As Long, i_editHeight As Long, i_memoWidth As Long, i_memoHeight As Long

Global i_labelStyle As Long, i_labelExtend As Long, i_buttonStyle As Long, i_buttonExtend As Long
Global i_CheckStyle As Long, i_checkExtend As Long, i_radioStyle As Long, i_radioExtend As Long
Global i_listStyle As Long, i_listExtend As Long, i_multiStyle As Long, i_multiExtend As Long
Global i_editStyle As Long, i_editExtend As Long, i_memoStyle As Long, i_memoExtend As Long

Global i_statusWidth As Long, i_statusHeight As Long
Global i_statusStyle As Long, i_statusExtend As Long

Global s_appPath As Asciiz * %MAX_PATH, s_find As String
Global s_inputIni As Asciiz * %MAX_PATH, s_outputIni As Asciiz * %MAX_PATH
Global s_inputTxt As Asciiz * %MAX_PATH, s_helpTxt As Asciiz * %MAX_PATH, s_OutputTxt As Asciiz * %MAX_PATH
Global i_ctl As Long, i_ctlCount As Long, i_section As Long, i_sectionCount As Long, s_section As Asciiz * 81, s_sectionList As String
Global i_style As Long, i_band As Long, i_bandCount As Long, s_control As Asciiz * 11, s_align As Asciiz * 11

FUNCTION FileToString(BYVAL s_file AS ASCIIZ * 256) AS STRING
LOCAL i_size AS LONG, h_file AS LONG, s_return AS STRING

IF LEN(DIR$(s_file, 7)) =0 THEN
s_return =""
ELSE
h_file =FREEFILE
OPEN s_file FOR BINARY AS h_file
i_size =LOF(h_file)
GET$ h_file, i_size, s_return
CLOSE h_file
END IF
FUNCTION =s_return
END FUNCTION

Function Id2Ctl(i_id As Long) As Long
Local I_return As Long
Array Scan a_id(), =Format$(i_id), To i_return
Decr i_return
Function =i_return
End Function

Function LabelSet(ByVal i_ctl As Long) As Long
If InStr("|" +a_misc(i_ctl) +"|", "|nolabel|") Then Exit Function

i =i_ctl -1
a_name(i) ="Label_" +a_name(i_ctl)
a_band(i) =a_band(i_ctl)
a_caption(i) =a_caption(i_ctl) +":"
a_control(i) ="label"
a_id(i) =Format$(Val(a_id(i_ctl)) -1)
a_left(i) =a_left(i_ctl)
a_top(i) =a_top(i_ctl)
a_style(i) =Format$(i_style Or i_labelStyle)
a_extend(i) =Format$(i_labelExtend)
x =Val(a_left(i)) +Val(a_width(i)) +i_labelPad
a_left(i_ctl) =Format$(x)
a_align(i) =a_align(i_ctl)
a_align(i_ctl) ="r"
End Function

Function CenterDialogOnDesktop(hDlg As Dword) As Long
Local Rct As RECT
Local dX As Long, dY As Long

Dialog Get Size hDlg To dX, dY
SystemParametersInfo %SPI_GETWORKAREA, 0, ByVal VarPtr(Rct), 0
Dialog Pixels hDlg, Rct.nLeft,  Rct.nTop    To Units Rct.nLeft,  Rct.nTop
Dialog Pixels hDlg, Rct.nRight, Rct.nBottom To Units Rct.nRight, Rct.nBottom
Dialog Set Loc hDlg, Rct.nLeft + (Rct.nRight - Rct.nLeft - dx) / 2, Rct.nTop + (Rct.nBottom - Rct.nTop - dy) / 2
End Function

Function HelpActivate(ByVal i_ctl As Long) As Long
Local i_start As Long, i_end As Long, i_len As Long, s_body As String

s_body =a_help(i_ctl)
If Len(s_body) =0 Then
s_body =GetSection(FileToString(s_helpTxt), a_name(i_ctl), "")
If Len(s_body) =0 Then s_body =a_tip(i_ctl)
Else
Replace "|" With $CRLF In s_body
End If
If Len(s_body) =0 Then
Local s as string
s ="No help defined for this dialog!"
Else
s =a_name(i_ctl)
s =GetSection(s_body, s, "No help defined for " +s +"!")
End If

if 0 Then
Local s_tempIni, s_tempTxt As ASCIIZ * %MAX_PATH
s_tempIni =s_appPath +"help_input.ini"
s_tempTxt =s_appPath +"help_input.txt"
If IsFile(s_tempIni) Then Kill s_tempIni
If IsFile(s_tempTxt) Then Kill s_tempTxt

s ="Help for " +a_name(i_ctl)
Ini_SetKey(s_tempIni, s, "control", "form")
'Ini_SetKey(s_tempIni, s, "align", Format$(h_dlg))
Ini_SetKey(s_tempIni, a_name(i_ctl), "control", "memo")
Ini_SetKey(s_tempIni, a_name(i_ctl), "misc", "ReadOnly|NoLabel")
Ini_SetKey(s_tempIni, a_name(i_ctl), "tip", "Press Escape to close")
Ini_SetKey(s_tempIni, "Close", "control", "button")
Ini_SetKey(s_tempIni, "Close", "id", Format$(%ID_CANCEL))
String2File("[[" +a_name(i_ctl) +"]]" +$CRLF +s_body, s_tempTxt)
' oIniForm.RunForm("help")
end if ' 0

s ="Help for " +a_name(i_ctl)
DialogShow(s, s_body)
End Function

Function DialogLoadFromIni() As Long
ini_DeleteSection(s_inputIni, "Results")
s_sectionList =ini_GetSectionsList(s_inputIni)
i_sectionCount =ParseCount(s_sectionList, $CrLf)
i_ctlCount =2 *(i_sectionCount -1) +1

'dimension arrays
Dim a_dlgBand(i_ctlCount), a_band(i_ctlCount), a_name(i_ctlCount), a_id(i_ctlCount), a_control(i_ctlCount), a_type(i_ctlCount), a_mask(i_ctlCount)
Dim a_caption(i_ctlCount), a_value(i_ctlCount), a_range(i_ctlCount), a_select(i_ctlCount), a_focus(i_ctlCount), a_result(i_ctlCount)
Dim a_align(i_ctlCount), a_left(i_ctlCount), a_top(i_ctlCount), a_width(i_ctlCount), a_height(i_ctlCount)
Dim a_fore(i_ctlCount), a_back(i_ctlCount), a_style(i_ctlCount), a_extend(i_ctlCount), a_tip(i_ctlCount), a_help(i_ctlCount), a_misc(i_ctlCount)

s_section =Parse$(s_sectionList, $CRLF, 1)
s_dlgInput =LCase$(ini_GetKey(s_inputIni, s_section, "input", "data"))
s_dlgOutput =LCase$(ini_GetKey(s_inputIni, s_section, "output", "data"))

'default control parameters
s_dlgFont =ini_GetKey(s_inputIni, s_section, "font", "MS Sans Serif")
i_dlgPoint =Val(ini_GetKey(s_inputIni, s_section, "point", "8"))
i_borderPad =Val(ini_GetKey(s_inputIni, s_section, "BorderPad", "7"))
i_labelPad =Val(ini_GetKey(s_inputIni, s_section, "LabelPad", "4"))

i_labelWidth =Val(ini_GetKey(s_inputIni, s_section, "LabelWidth", "40"))
i_labelHeight =Val(ini_GetKey(s_inputIni, s_section, "LabelHeight", "8"))
i_labelStyle =Val(ini_GetKey(s_inputIni, s_section, "LabelStyle", Format$(%SS_RIGHT Or %SS_CENTERIMAGE)))
i_labelExtend =Val(ini_GetKey(s_inputIni, s_section, "LabelExtend", Format$(%WS_EX_LEFT)))

i_buttonWidth =Val(ini_GetKey(s_inputIni, s_section, "ButtonWidth", "50"))
i_buttonHeight =Val(ini_GetKey(s_inputIni, s_section, "ButtonHeight", "14"))
i_buttonStyle =Val(ini_GetKey(s_inputIni, s_section, "ButtonStyle", Format$(%BS_CENTER Or %BS_VCENTER Or %WS_TABSTOP Or %BS_NOTIFY)))
i_buttonExtend =Val(ini_GetKey(s_inputIni, s_section, "ButtonExtend", Format$(%WS_EX_LEFT)))

i_checkWidth =Val(ini_GetKey(s_inputIni, s_section, "CheckWidth", "40"))
i_checkHeight =Val(ini_GetKey(s_inputIni, s_section, "CheckHeight", "14"))
i_checkStyle =Val(ini_GetKey(s_inputIni, s_section, "CheckStyle", Format$(%BS_CENTER Or %BS_VCENTER Or %WS_TABSTOP Or %BS_NOTIFY)))
i_checkExtend =Val(ini_GetKey(s_inputIni, s_section, "CheckExtend", Format$(%WS_EX_LEFT)))

i_radioWidth =Val(ini_GetKey(s_inputIni, s_section, "RadioWidth", "40"))
i_radioHeight =Val(ini_GetKey(s_inputIni, s_section, "RadioHeight", "14"))
i_radioStyle =Val(ini_GetKey(s_inputIni, s_section, "RadioStyle", Format$(%BS_CENTER Or %BS_VCENTER Or %BS_NOTIFY)))
i_radioExtend =Val(ini_GetKey(s_inputIni, s_section, "RadioExtend", Format$(%WS_EX_LEFT)))

i_listWidth =Val(ini_GetKey(s_inputIni, s_section, "ListWidth", "100"))
i_listHeight =Val(ini_GetKey(s_inputIni, s_section, "ListHeight", "50"))
i_listStyle =Val(ini_GetKey(s_inputIni, s_section, "ListStyle", Format$(%WS_TABSTOP Or %WS_VSCROLL Or %LBS_NOTIFY)))
i_listExtend =Val(ini_GetKey(s_inputIni, s_section, "ListExtend", Format$(%WS_EX_CLIENTEDGE Or %WS_EX_LEFT)))

i_multiWidth =Val(ini_GetKey(s_inputIni, s_section, "MultiWidth", "100"))
i_multiHeight =Val(ini_GetKey(s_inputIni, s_section, "MultiHeight", "50"))
i_multiStyle =Val(ini_GetKey(s_inputIni, s_section, "MultiStyle", Format$(%LBS_MULTIPLESEL Or %WS_TABSTOP Or %WS_VSCROLL Or %LBS_NOTIFY)))
i_multiExtend =Val(ini_GetKey(s_inputIni, s_section, "MultiExtend", Format$(%WS_EX_CLIENTEDGE Or %WS_EX_LEFT)))

i_editWidth =Val(ini_GetKey(s_inputIni, s_section, "EditWidth", "100"))
i_editHeight =Val(ini_GetKey(s_inputIni, s_section, "EditHeight", "12"))
i_editStyle =Val(ini_GetKey(s_inputIni, s_section, "EditStyle", Format$(%WS_TABSTOP Or %WS_BORDER Or %ES_LEFT Or %ES_AUTOHSCROLL)))
i_editExtend =Val(ini_GetKey(s_inputIni, s_section, "EditExtend", Format$(%WS_EX_CLIENTEDGE Or %WS_EX_LEFT)))

i_memoWidth =Val(ini_GetKey(s_inputIni, s_section, "MemoWidth", "100"))
i_memoHeight =Val(ini_GetKey(s_inputIni, s_section, "MemoHeight", "50"))
i_memoStyle =Val(ini_GetKey(s_inputIni, s_section, "MemoStyle", Format$(%WS_TABSTOP Or %WS_BORDER Or %ES_LEFT Or %WS_VSCROLL Or %ES_MULTILINE Or %ES_WANTRETURN)))
i_memoExtend =Val(ini_GetKey(s_inputIni, s_section, "MemoExtend", Format$(%WS_EX_CLIENTEDGE Or %WS_EX_LEFT)))

i_statusWidth =Val(ini_GetKey(s_inputIni, s_section, "statusWidth", "40"))
i_statusHeight =Val(ini_GetKey(s_inputIni, s_section, "statusHeight", "12"))
i_statusStyle =Val(ini_GetKey(s_inputIni, s_section, "statusStyle", Format$(%WS_CHILD Or %WS_VISIBLE)))
i_statusExtend =Val(ini_GetKey(s_inputIni, s_section, "statusExtend", Format$(%WS_EX_TRANSPARENT Or %WS_EX_LEFT Or %WS_EX_LTRREADING Or %WS_EX_RIGHTSCROLLBAR)))

Local i_section as long
For i_section =1 To i_SectionCount
s_section =Parse$(s_sectionList, $CrLf, i_section)
i_ctl =2 *(i_section -1)
a_name(i_ctl) =s_section
a_caption(i_ctl) =ini_GetKey(s_inputIni, s_section, "caption", s_section)
a_control(i_ctl) =LCase$(ini_GetKey(s_inputIni, s_section, "control", IIf$(i_ctl, "edit", "form")))
a_id(i_ctl) =ini_GetKey(s_inputIni, s_section, "id", Format$(ControlSetID(i_ctl)))
a_type(i_ctl) =ini_GetKey(s_inputIni, s_section, "type", "")
a_mask(i_ctl) =ini_GetKey(s_inputIni, s_section, "mask", "")
a_range(i_ctl) =ini_GetKey(s_inputIni, s_section, "range", "")
a_value(i_ctl) =ini_GetKey(s_inputIni, s_section, "value", "")
a_select(i_ctl) =ini_GetKey(s_inputIni, s_section, "selection", "")
' a_select(i_ctl) =ini_GetKey(s_inputIni, s_section, "select", "")
a_focus(i_ctl) =ini_GetKey(s_inputIni, s_section, "focus", "")
a_align(i_ctl) =LCase$(ini_GetKey(s_inputIni, s_section, "align", ""))
a_left(i_ctl) =ini_GetKey(s_inputIni, s_section, "left", IIf$(i_ctl, "", "-1"))
a_top(i_ctl) =ini_GetKey(s_inputIni, s_section, "top", IIf$(i_ctl, "", "-1"))
a_width(i_ctl) =ini_GetKey(s_inputIni, s_section, "width", "")
a_height(i_ctl) =ini_GetKey(s_inputIni, s_section, "height", IIf$(i_ctl, "", "250"))
a_fore(i_ctl) =ini_GetKey(s_inputIni, s_section, "fore", IIf$(i_ctl, a_fore(0), ""))
a_back(i_ctl) =ini_GetKey(s_inputIni, s_section, "back", IIf$(i_ctl, a_back(0), ""))
a_style(i_ctl) =ini_GetKey(s_inputIni, s_section, "style", "")
a_extend(i_ctl) =ini_GetKey(s_inputIni, s_section, "extend", Format$(ControlSetExtend(i_ctl)))
a_tip(i_ctl) =ini_GetKey(s_inputIni, s_section, "tip", "")
a_help(i_ctl) =ini_GetKey(s_inputIni, s_section, "help", "")
a_misc(i_ctl) =LCASE$(ini_GetKey(s_inputIni, s_section, "misc", ""))
Next i_section

Dialog Font s_dlgFont, i_dlgPoint
i_dlgWidth =Val(a_width(0))
i_dlgHeight =Val(a_height(0))
s_section =a_name(0)
a_style(0) =ini_GetKey(s_inputIni, s_section, "style", Format$(%DS_3DLOOK Or %DS_SETFONT Or %DS_MODALFRAME Or %DS_NOFAILCREATE Or %WS_BORDER Or %WS_CLIPSIBLINGS Or %WS_DLGFRAME Or %WS_POPUP Or %WS_SYSMENU))
a_extend(0) =ini_GetKey(s_inputIni, s_section, "extend", Format$(%WS_EX_LEFT Or %WS_EX_LTRREADING Or %WS_EX_RIGHTSCROLLBAR))
End Function

Function DialogActivate() As Long
Local h_parent as DWord

h_parent =Val(a_align(0))
If h_parent =0 Or h_parent =100 Then
h_parent =%HWND_DESKTOP
Else
a_style(0) =Format$(Val(a_style(0)) Or %WS_CHILD Or %WS_VISIBLE)
End If

Local s as String
' msgbox a_control(0)
Select Case As Const$ a_control(0)
Case "DialogShow"
' DialogShow(a_caption(0), a_value(0))
MsgBox a_value(0), 0, a_caption(0)

Case "DialogConfirm"
a_value(0) = DialogConfirm(ByCopy a_caption(0), ByCopy a_misc(0), ByCopy a_value(0))
Case "DialogOpenFolder"
' a_value(0) =BrowseForFolder(%HWND_DESKTOP, ByCopy a_caption(0), a_value(0), %FALSE)
Display Browse %HWND_DESKTOP, , , a_caption(0), a_value(0), %NULL to a_value(0)
ini_SetKey(s_outputIni, "Results", a_caption(0), a_value(0))
Case "DialogOpenFile", "DialogSaveFile"
If Len(a_range(0)) =0 Then a_range(0) =CurDir$
a_style(0) =""
If a_control(0) ="open" Then
If Len(a_style(0)) =0 Then a_style(0) =Format$(%OFN_FILEMUSTEXIST Or %OFN_HIDEReadOnly)
' i_result =OpenFileDialog(%HWND_DESKTOP, a_caption(0), a_value(0), a_range(0), a_misc(0), a_type(0), Val(a_style(0)))
Display OpenFile %HWND_DESKTOP, , , a_caption(0), "", "", a_value(0), PathName$(extn, a_value(0)), %NULL to a_value(0)
Else 
' msgbox format$(InStr(a_misc(0), "nooverwriteprompt"))
If IsFalse InStr(LCase$(a_misc(0)), LCase$("NoOverWritePrompt")) And Len(a_style(0)) =0 Then a_style(0) =Format$(%OFN_OVERWRITEPROMPT)
' i_result =SaveFileDialog(%HWND_DESKTOP, a_caption(0), a_value(0), a_range(0), a_misc(0), a_type(0), Val(a_style(0)))
Display SaveFile %HWND_DESKTOP, , , a_caption(0), "", "", a_value(0), PathName$(EXTN, a_value(0)), %NULL to a_value(0)
' msgbox a_value(0)
End If
If i_result Then
If Dir$(s_outputIni, 7) <>"" Then Kill s_outputIni
ini_SetKey(s_outputIni, "Results", a_caption(0), a_value(0))
End If
Case Else

Dialog New h_parent, a_caption(0), , , i_dlgWidth, i_dlgHeight, _
Val(a_style(0)), Val(a_extend(0)) To h_dlg

If a_Fore(0) <>"" Or a_back(0) <>"" Then Dialog Set Color h_dlg, Val(a_fore(0)), Val(a_back(0))

Local i_ctl As Long
For i_ctl =1 To i_ctlCount
If Len(a_name(i_ctl)) =0 Then Iterate

Select Case As Const$ a_control(i_ctl)
Case "label"
Control Add Label, h_dlg, Val(a_id(i_ctl)), a_caption(i_ctl), Val(a_left(i_ctl)), Val(a_top(i_ctl)), Val(a_width(i_ctl)), Val(a_height(i_ctl)), _
Val(a_style(i_ctl)), Val(a_extend(i_ctl)) Call LabelEvent

Case "button"
Control Add Button, h_dlg, Val(a_id(i_ctl)), a_caption(i_ctl), Val(a_left(i_ctl)), Val(a_top(i_ctl)), Val(a_width(i_ctl)), Val(a_height(i_ctl)), _
Val(a_style(i_ctl)), Val(a_extend(i_ctl)) Call ButtonEvent

Case "radio"
Control Add Option, h_dlg, Val(a_id(i_ctl)), a_caption(i_ctl), Val(a_left(i_ctl)), Val(a_top(i_ctl)), Val(a_width(i_ctl)), Val(a_height(i_ctl)), _
Val(a_style(i_ctl)), Val(a_extend(i_ctl)) Call RadioEvent
Control Set Check h_dlg, Val(a_id(i_ctl)), Val(a_value(i_ctl))

Case "check"
Control Add CheckBox, h_dlg, Val(a_id(i_ctl)), a_caption(i_ctl), Val(a_left(i_ctl)), Val(a_top(i_ctl)), Val(a_width(i_ctl)), Val(a_height(i_ctl)), _
Val(a_style(i_ctl)), Val(a_extend(i_ctl)) Call CheckEvent
Control Set Check h_dlg, Val(a_id(i_ctl)), Val(a_value(i_ctl))

Case "list"
s =a_range(i_ctl)
' msgbox format$(len(s))
If Len(s) =0 Then
s =GetSection(FileToString(s_inputTxt), a_name(i_ctl), "")
' msgbox s
Replace $CRLF With "|" In S
End If
i =ParseCount(s, "|")
ReDim a(i -1)
Parse s, a(), "|"
' DialogShow(a(0), a(1))
Control Add ListBox, h_dlg, Val(a_id(i_ctl)), a(), Val(a_left(i_ctl)), Val(a_top(i_ctl)), Val(a_width(i_ctl)), Val(a_height(i_ctl)), _
Val(a_style(i_ctl)), Val(a_extend(i_ctl)) Call ListEvent
' msgbox a_select(i_ctl)
ListBox Select h_dlg, Val(a_id(i_ctl)), Val(a_select(i_ctl))
Global i_lb As Long
i_lb = i_ctl
Global h_lb as DWord
CONTROL HANDLE h_Dlg, Val(a_id(i_ctl)) TO h_lb
'MsgBox Format$(h_ctl)
'PostMessage(h_lb, %LB_SETSEL, 0, 0)

Case "multi"
s =a_range(i_ctl)
If Len(s) =0 Then
' msgbox s_inputTxt
' MsgBox Format$(IsFile(s_inputTxt))
' MsgBox FileToString(s_inputTxt)
s =GetSection(FileToString(s_inputTxt), a_name(i_ctl), "")
' msgbox s
Replace $CRLF With "|" In S
End If
i =ParseCount(s, "|")
' MsgBox Format$(i)
ReDim a(i -1)
Parse s, a(), "|"
' MsgBox Format$(ArrayAttr(a(), 4))
Local j as Long
For j = 0 to UBound(a)
If Len(a(j)) = 0 Then a(j) = " "
Next j

Control Add ListBox, h_dlg, Val(a_id(i_ctl)), a(), Val(a_left(i_ctl)), Val(a_top(i_ctl)), Val(a_width(i_ctl)), Val(a_height(i_ctl)), _
Val(a_style(i_ctl)), Val(a_extend(i_ctl)) Call MultiEvent
s =Trim$(a_select(i_ctl))
If s <>"" Then
i_count =ParseCount(s, "|")
Local i As Long
For i =1 To i_count
Control Send h_dlg, Val(a_id(i_ctl)), %LB_SetSEL, 1, (Val(Parse$(s, "|", i)) -1)
Next i
End If

s =LCase$(Trim$(a_misc(i_ctl)))
If InStr(s, "unselectall") Then
ListBox Unselect h_dlg, Val(a_id(i_ctl)), 0
ElseIf InStr(s, "selectall") Then
ListBox Select h_dlg, Val(a_id(i_ctl)), 0
End If
Case "edit"
Control Add TextBox, h_dlg, Val(a_id(i_ctl)), a_value(i_ctl), Val(A_left(i_ctl)), Val(a_top(i_ctl)), Val(a_width(i_ctl)), Val(a_height(i_ctl)), _
Val(a_style(i_ctl)), Val(a_extend(i_ctl)) Call EditEvent
If Val(a_mask(i_ctl)) >0 Then Control Send h_dlg, Val(a_id(i_ctl)), %EM_SETLIMITTEXT, Val(a_mask(i_ctl)), 0
If (Val(a_style(i_ctl)) And %ES_ReadOnly) Then Control Post h_dlg, Val(a_id(i_ctl)), %EM_SETSEL, 0, 0

Case "memo"
If Len(a_value(i_ctl)) =0 Then
a_value(i_ctl) =GetSection(FileToString(s_inputTxt), a_name(i_ctl), "")
Else
Replace "|" With $CRLF In A_value(i_ctl)
End If

if 1 then
s =a_value(i_ctl)
a_value(i_ctl) =""
Control Add TextBox, h_dlg, Val(a_id(i_ctl)), a_value(i_ctl), Val(A_left(i_ctl)), Val(a_top(i_ctl)), Val(a_width(i_ctl)), Val(a_height(i_ctl)), _
Val(a_style(i_ctl)), Val(a_extend(i_ctl)) Call MemoEvent
a_value(i_ctl) =s
CONTROL SEND h_dlg, Val(a_id(i_ctl)), %EM_EXLIMITTEXT, 0, 2000000
CONTROL SEND h_dlg, Val(a_id(i_ctl)), %WM_SETTEXT, 0, STRPTR(a_value(i_ctl))

If Val(a_mask(i_ctl)) >0 Then Control Send h_dlg, Val(a_id(i_ctl)), %EM_SETLIMITTEXT, Val(a_mask(i_ctl)), 0
else
a_style(i_ctl) =Format$(Val(a_style(i_ctl)) Or %WS_CHILD Or %WS_VISIBLE)
'InitCommonControls
LoadLibrary("RICHED32.DLL")
s =""
Control Add "richedit", h_dlg, Val(a_id(i_ctl)), s, _
Val(a_left(i_ctl)), Val(a_top(i_ctl)), Val(a_width(i_ctl)), Val(a_height(i_ctl)), _
Val(a_style(i_ctl)), Val(a_extend(i_ctl))

CONTROL SEND h_dlg, Val(a_id(i_ctl)), %EM_EXLIMITTEXT, 0, 2000000
CONTROL SEND h_dlg, Val(a_id(i_ctl)), %WM_SETTEXT, 0, STRPTR(a_value(i_ctl))
end if
If (Val(a_style(i_ctl)) And %ES_ReadOnly) Then Control Post h_dlg, Val(a_id(i_ctl)), %EM_SETSEL, 0, 0
Case "status"
Local icc As INIT_COMMON_CONTROLSEX
icc.dwICC =%ICC_BAR_CLASSES
icc.dwSize =SizeOf(icc)
' InitCommonControlsEx(icc)
' Control Add "msctls_statusbar32", h_dlg, Val(a_id(i_ctl)), a_tip(i_ctl), Val(a_left(i_ctl)), Val(a_top(i_ctl)), Val(a_width(i_ctl)), Val(a_height(i_ctl)), Val(a_style(i_ctl)), Val(a_extend(i_ctl))
' Control Add StatusBar, h_dlg, val(a_id(i_ctl)), a_tip(i_ctl), val(a_left(i_ctl)), val(a_top(i_ctl)), val(a_width(i_ctl)), val(a_height(i_ctl)), val(a_style(i_ctl)), val(a_extend(i_ctl))
End Select

If a_fore(i_ctl) <>a_fore(0) Or a_back(i_ctl) <>a_back(0) Then Control Set Color h_dlg, Val(a_id(i_ctl)), Val(a_fore(i_ctl)), Val(a_back(i_ctl))
Next i_ctl

'Dialog Show State h_dlg, %WS_MAXIMIZE
'CenterDialogOnDesktop(h_dlg)
Dialog Get Loc h_dlg To i_dlgLeft, i_dlgTop
'Dialog Get size h_dlg To i_dlgWidth, i_dlgHeight
Dialog Get Client h_dlg To i_dlgWidth, i_dlgHeight
a_left(0) =Format$(i_dlgLeft)
a_top(0) =Format$(i_dlgTop)
a_width(0) =Format$(i_dlgWidth)
a_height(0) =Format$(i_dlgHeight)

DialogSetHotKeys()
If 1 Then
Dialog Show Modal h_dlg, Call DialogEvent To i_result
Else
Dialog Show Modeless h_dlg, Call DialogEvent
i_result =1
While i_result
Dialog DoEvents To i_result
WEnd
End If
Function =i_result
End Select
End Function

Function DialogGroupControls() As Long
Local i_width As Long

s_control =""
i_band =1

Local i_ctl As Long
For i_ctl =1 To i_ctlCount
If Len(a_name(i_ctl)) =0 Then Iterate

s_align =Trim$(Mid$(a_align(i_ctl) +Space$(1), 1, 1))
s_align =IIf$(InStr("rd", s_align), s_align, "")
If s_align ="" Then s_align =IIf$(A_control(i_ctl) =s_control And IsFalse InStr("list|multi|edit|memo", a_control(i_ctl)), "r", "d")
a_align(i_ctl) =s_align
If (i_ctl >2 And a_align(i_ctl) ="d") Then Incr i_band
Incr a_dlgBand(i_band).Count
a_band(i_ctl) =i_band
i_width =4 *Len(a_caption(i_ctl)) +4
i_statusWidth =Max(i_statusWidth, Len(a_tip(i_ctl)) +8)

Select Case As Const$ a_control(i_ctl)
Case "label"

Case "button"
a_dlgBand(i_band).ButtonWidth =Max(a_dlgBand(i_ctl).ButtonWidth, i_width)

Case "check"
a_dlgBand(i_band).CheckWidth =Max(a_dlgBand(i_ctl).CheckWidth, i_width)

Case "radio"
a_dlgBand(i_band).RadioWidth =Max(a_dlgBand(i_ctl).RadioWidth, i_width)

Case "list","multi","edit","memo"
i_labelWidth =Max(i_labelWidth, i_width +4)
End Select
s_control =a_control(i_ctl)
Next i_ctl

If IsFalse InStr("|" +a_misc(0) +"|", "nostatus") Then
Incr i_band
a_name(i_ctlCount) ="status"
a_control(i_ctlCount) ="status"
a_id(i_ctlCount) =Format$(%ID_STATUSBAR)
a_align(i_ctlCount) ="d"
' a_tip(i_ctlCount) ="Ready"
a_tip(i_ctlCount) ="Ready"
a_band(i_ctlCount) =i_band
a_dlgBand(i_band).Count =1
'a_dlgBand(i_band).Height =a_height(i_ctlCount)
End If
i_bandCount =i_band
End Function

Function DialogSizeControls() As Long
Local i_width As Long, i_height As Long

Local i_ctl As Long
For i_ctl =1 To i_ctlCount
If Len(a_name(i_ctl)) =0 Then Iterate

i_band =a_band(i_ctl)
i_width =4 *Len(a_caption(i_ctl)) +4

Select Case As Const$ a_control(i_ctl)
Case "label"
a_width(i_ctl) =Format$(i_width)
a_height(i_ctl) =Format$(i_labelHeight)

Case "button"
a_width(i_ctl) =Format$(a_dlgBand(i_band).ButtonWidth)
a_height(i_ctl) =Format$(i_buttonHeight)

Case "check"
a_width(i_ctl) =Format$(a_dlgBand(i_band).CheckWidth)
a_height(i_ctl) =Format$(i_checkHeight)

Case "radio"
a_width(i_ctl) =Format$(a_dlgBand(i_band).RadioWidth)
a_height(i_ctl) =Format$(i_radioHeight)

Case "list","multi","edit","memo"
If IsFalse InStr("|" +a_misc(i_ctl) +"|", "|nolabel|") Then
i =i_ctl -1
a_width(i) =Format$(i_labelWidth)
a_height(i) =Format$(IIf(InStr("list|multi", a_control(i_ctl)), 10, 12))
a_dlgBand(i_band).CtlWidth =a_dlgBand(i_band).CtlWidth +Val(a_width(i)) +i_LabelPad
a_dlgBand(i_band).DlgWidth =a_dlgBand(i_band).DlgWidth +Val(a_width(i)) +i_LabelPad
End If

Select Case As Const$ a_control(i_ctl)
Case "list","multi"
local s as string
s =a_range(i_ctl)
If Len(s) =0 Then
a_width(i_ctl) =Format$(Iif(a_control(i_ctl) ="list", i_listWidth, i_multiWidth))
Else
Local i as Long
For i =1 To ParseCount(s, "|")
a_width(i_ctl) =Format$(Max(Val(a_width(i_ctl)), 8 +4 *Len(Parse$(s, "|", i))))
Next i
End If
a_height(i_ctl) =Format$(IIf(a_control(i_ctl) ="list", i_listHeight, i_multiHeight))

Case "edit","memo"
i =Val(a_mask(i_ctl))
a_width(i_ctl) =Format$(IIf(i =0, Iif(a_control(i_ctl) ="edit", i_editWidth, i_memoWidth), 4 *i +4))
a_height(i_ctl) =Format$(IIf(a_control(i_ctl) ="edit", i_editHeight, i_memoHeight))
End Select
Case "status"
a_width(i_ctl) =Format$(i_statusWidth)
a_height(i_ctl) =Format$(i_statusHeight)
End Select

a_dlgBand(i_band).Height =Max(a_dlgBand(i_band).Height, Val(a_height(i_ctl)))
a_dlgBand(i_band).CtlWidth =a_dlgBand(i_band).CtlWidth +Val(a_width(i_ctl))
a_dlgBand(i_band).DlgWidth =a_dlgBand(i_band).DlgWidth +Val(a_width(i_ctl)) +i_borderPad
i_dlgWidth =Max(i_dlgWidth, a_dlgBand(i_band).DlgWidth)
Next i_ctl
i_dlgWidth =i_dlgWidth +i_borderPad
If a_control(i_ctlCount) ="status" Then a_width(i_ctlCount) =Format$(i_dlgWidth)
End Function

Function DialogPositionControls() As Long
Local i_band as Long
For i_band =1 To i_bandCount
a_dlgBand(i_band).Top =a_dlgBand(i_band -1).Bottom +Iif(i_band =i_bandCount And a_control(i_ctlCount) ="status", 2 *i_borderPad, i_borderPad)
a_dlgBand(i_band).Bottom =a_dlgBand(i_band).Top +a_dlgBand(i_band).Height
i_dlgHeight =a_dlgBand(i_band).Bottom
xSpace =(i_dlgWidth -a_dlgBand(i_band).CtlWidth)/(a_dlgBand(i_band).Count +1)

'x =i_borderPad
x =xSpace
s_control =""

Local i_ctl As Long
For i_ctl =1 To i_ctlCount
If a_band(i_ctl) <>i_band Then Iterate

i_style =IIf(a_control(i_ctl) =s_control, 0, %WS_GROUP)

a_left(i_ctl) =Format$(x)
a_top(i_ctl) =Format$(a_dlgBand(i_band).Top)

Select Case As Const$ a_control(i_ctl)
Case "label"
If i_style =0 Then i_style =i_labelStyle
Case "button"
If Len(a_style(i_ctl)) =0 Then
If InStr("|" +a_misc(i_ctl) +"|", "|default|") Or Val(a_id(i_ctl)) =%ID_OK Then i_style =i_style Or %BS_DEFAULT
a_style(i_ctl) =Format$(i_style Or i_buttonStyle)
End If

Case "radio"
If a_control(i_ctl) =s_control Then
i_style =0
Else
i_style =i_style Or %WS_GROUP Or %WS_TABSTOP
i =i_ctl +2
i_count =0 'number of checked radio buttons
While a_control(i) ="radio" And i <=i_ctlCount
i =i +2
i_count =Max(i_count, Val(a_value(i)))
Wend
If i_count =0 Then a_value(i_ctl) ="1"
End If
If Len(a_style(i_ctl)) =0 Then a_style(i_ctl) =Format$(i_style Or i_radiostyle)

Case "check"
If Len(a_style(i_ctl)) =0 Then a_style(i_ctl) =Format$(i_style Or i_checkStyle)

Case "list","multi"
LabelSet(i_ctl)
If Len(a_style(i_ctl)) =0 Then
a_style(i_ctl) =Format$(i_style Or IIf(a_control(i_ctl) ="list", i_listStyle, i_multiStyle))
If InStr("|" +a_misc(i_ctl) +"|", "|sort|") Then a_style(i_ctl) =Format$(Val(a_style(i_ctl)) Or %LBS_SORT)
End If

Case "edit","memo"
LabelSet(i_ctl)
If Len(a_style(i_ctl)) =0 Then
a_style(i_ctl) =Format$(i_style Or IIf(a_control(i_ctl) ="edit", i_editStyle, i_memoStyle))
If InStr("|" +a_misc(i_ctl) +"|", "|readonly|") Then a_style(i_ctl) =Format$(Val(a_style(i_ctl)) Or %ES_ReadOnly Or %ES_SAVESEL)
If IsTrue InStr(a_caption(i_ctl), "Password") Or IsTrue InStr("|" +a_misc(i_ctl) +"|", "|password|") Then a_style(i_ctl) =Format$(Val(a_style(i_ctl)) Or %ES_PASSWORD)
End If

Case "status"
If Len(a_style(i_ctl)) =0 Then a_style(i_ctl) =Format$(i_statusStyle)
If Len(a_extend(i_ctl)) =0 Then a_extend(i_ctl) =Format$(i_statusExtend)
End Select

x =x +Val(a_width(i_ctl)) +xSpace
s_control =a_control(i_ctl)
Next i_ctl
Next i_band
If a_control(i_ctlCount) ="status" Then a_left(i_ctlCount) ="0"
i_dlgHeight =i_dlgHeight +i_borderPad
End Function

Function DialogSaveToIni(ByVal i_button As Long) As Long
a_result(0) =A_caption(Id2Ctl(i_id))
If Dir$(s_outputIni, 7) <>"" Then Kill s_outputIni
If Dir$(s_outputTxt, 7) <>"" Then Kill s_outputTxt
ini_SetKey(s_outputIni, "Results", "dummy", " ")
Ini_DeleteKey(s_outputIni, "Results", "dummy")
'i_count =gettickcount()
Local i_ctl as Long
For i_ctl =0 To i_ctlCount
If Len(a_name(i_ctl)) =0 Then Iterate
If a_caption(i_ctl) =a_name(i_ctl) Then a_caption(i_ctl) =""
If a_extend(i_ctl) ="0" Then a_extend(i_ctl) =""
i_id =Val(a_id(i_ctl))
Select Case As Const$ a_control(i_ctl)
Case "button"
i =IIf(i_ctl =i_button, 1, 0)
local s as string
s =Iif$(i, Format$(i), "")
Case "check","radio"
Control Get Check h_dlg, i_id To i
s =Format$(i)
Case "edit"
Control Get Text h_dlg, i_id To s
a_value(i_ctl) =s
Case "multi"
If Len(a_range(i_ctl) )=0 Then
s =GetSection(FileToString(s_inputTxt), a_name(i_ctl), "")
Append2File("[[" +a_name(i_ctl) +"]]" +$CRLF +s, s_outputTxt)
End If

a_select(i_ctl) =""
' Control Send h_dlg, i_id, %LB_GETSELCOUNT, 0, 0 To i_count
ListBox Get SelCount h_dlg, i_id To i_count
If i_count Then
Dim a_multi(i_count -1)
' Control Send h_dlg, i_id, %LB_GETSELITEMS, i_count, VarPtr(a_multi(0))
s_list =""
Local i as Long
' For i =0 To i_count -1
Local b_loop As Long
i = 1
b_Loop = %True
Local i_index As Long
While b_loop
ListBox Get Select h_dlg, i_id, i To i_index
' msgbox format$(i_index)
If i_index = 0 Then Exit Do
ListBox Get Text h_dlg, i_id, i_index To s_item
s_list =s_list +s_item +"|"
' a_select(i_ctl) =a_select(i_ctl) +Format$(i_index) +"|"
a_select(i_ctl) =a_select(i_ctl) +s_item +"|"
i = i_index + 1
WEnd
' Control Send h_dlg, i_id, %LB_GETTEXT, a_multi(i), VarPtr(c)
' s_list =s_list +c +"|"
' a_select(i_ctl) =a_select(i_ctl) +Format$(a_multi(i) +1) +"|"
' Next
Else
s =""
End If
s =Left$(s_list, (Len(s_list) -1))
a_value(i_ctl) = s
a_select(i_ctl) =Left$(a_select(i_ctl), Len(a_select(i_ctl)) -1)
s = a_select(i_ctl)

Case "list"
If Len(a_range(i_ctl) )=0 Then
s =GetSection(FileToString(s_inputTxt), a_name(i_ctl), "")
Append2File("[[" +a_name(i_ctl) +"]]" +$CRLF +s, s_outputTxt)
End If

ListBox Get Text h_dlg, i_id To s
a_value(i_ctl) =s
Control Send h_dlg, i_id, %LB_GETCURSEL, 0, 0 To i
a_select(i_ctl) =Format$(i +1)

Case "memo"
Control Get Text h_dlg, i_id To s
Append2File("[[" +a_name(i_ctl) +"]]" +$CRLF +s, s_outputTxt)
a_value(i_ctl) =""
s =""

Case Else
s =""
End Select
ini_SetKey(s_outputIni, "Results", a_name(i_ctl), s)
If s_dlgOutput ="all" Then
ControlSaveSettings(i_ctl)
End If
Next i_ctl
If s_dlgOutput ="all" Then DialogSaveSettings()
End Function

Function DialogSetHotKeys() As Dword
Dim a_keys(11)
Local h_return As Dword
i =0
a_keys(i).fvirt =%FVIRTKEY
a_keys(i).key =%VK_F1
a_keys(i).cmd =%ID_F1
i =1
a_keys(i).fvirt =%FVIRTKEY
a_keys(i).key =%VK_F3
a_keys(i).cmd =%ID_F3

i =2
a_keys(i).fvirt =%FSHIFT Or %FVIRTKEY
a_keys(i).key =%VK_F3
a_keys(i).cmd =%ID_SHIFT_F3
i =3
a_keys(i).fvirt =%FVIRTKEY
a_keys(i).key =%VK_F5
a_keys(i).cmd =%ID_F5
i =4
a_keys(i).fvirt =%FCONTROL Or %FVIRTKEY
a_keys(i).key =%VK_R
a_keys(i).cmd =%ID_CONTROL_R
i =5
a_keys(i).fvirt =%FCONTROL Or %FVIRTKEY
a_keys(i).key =%VK_s
a_keys(i).cmd =%ID_CONTROL_S
i =6
a_keys(i).fvirt =%FCONTROL Or %FVIRTKEY
a_keys(i).key =%VK_RETURN
a_keys(i).cmd =%ID_CONTROL_ENTER
i =7
a_keys(i).fvirt =%FCONTROL Or %FVIRTKEY
a_keys(i).key =%VK_F
a_keys(i).cmd =%ID_CONTROL_F
i =8
a_keys(i).fvirt =%FCONTROL Or %FSHIFT Or %FVIRTKEY
a_keys(i).key =%VK_F
a_keys(i).cmd =%ID_CONTROL_SHIFT_F
i =9
a_keys(i).fvirt =%FCONTROL Or %FVIRTKEY
a_keys(i).key =%VK_A
a_keys(i).cmd =%ID_CONTROL_A
i =10
a_keys(i).fvirt =%FCONTROL Or %FSHIFT Or %FVIRTKEY
a_keys(i).key =%VK_A
a_keys(i).cmd =%ID_CONTROL_SHIFT_A
Accel Attach h_dlg, a_keys() To h_return
Function =h_return
End Function

Function DialogSaveSettings() As Long
s_section =a_name(0)
s_dlgInput ="all"
Ini_SetKey(s_outputIni, s_section, "input", s_dlgInput)
s_dlgOutput ="data"
Ini_SetKey(s_outputIni, s_section, "output", s_dlgOutput)
Ini_SetKey(s_outputIni, s_section, "font", s_dlgFont)
Ini_SetKey(s_outputIni, s_section, "point", Format$(i_dlgPoint))
Ini_SetKey(s_outputIni, s_section, "LabelPad", Format$(i_labelPad))
Ini_SetKey(s_outputIni, s_section, "BorderPad", Format$(i_borderPad))

Ini_SetKey(s_outputIni, s_section, "LabelWidth", Format$(i_labelWidth))
Ini_SetKey(s_outputIni, s_section, "LabelHeight", Format$(i_labelHeight))
Ini_SetKey(s_outputIni, s_section, "LabelStyle", Format$(i_labelStyle))
Ini_SetKey(s_outputIni, s_section, "LabelExtend", Iif$(i_labelExtend =0, "", Format$(i_labelExtend)))
Ini_SetKey(s_outputIni, s_section, "ButtonWidth", Format$(i_ButtonWidth))
Ini_SetKey(s_outputIni, s_section, "ButtonHeight", Format$(i_ButtonHeight))
Ini_SetKey(s_outputIni, s_section, "ButtonStyle", Format$(i_ButtonStyle))
Ini_SetKey(s_outputIni, s_section, "ButtonExtend", Iif$(i_buttonExtend =0, "", Format$(i_ButtonExtend)))

Ini_SetKey(s_outputIni, s_section, "CheckWidth", Format$(i_CheckWidth))
Ini_SetKey(s_outputIni, s_section, "CheckHeight", Format$(i_CheckHeight))
Ini_SetKey(s_outputIni, s_section, "CheckStyle", Format$(i_CheckStyle))
Ini_SetKey(s_outputIni, s_section, "CheckExtend", Iif$(i_checkExtend =0, "", Format$(i_CheckExtend)))

Ini_SetKey(s_outputIni, s_section, "RadioWidth", Format$(i_RadioWidth))
Ini_SetKey(s_outputIni, s_section, "RadioHeight", Format$(i_RadioHeight))
Ini_SetKey(s_outputIni, s_section, "RadioStyle", Format$(i_RadioStyle))
Ini_SetKey(s_outputIni, s_section, "RadioExtend", Iif$(i_radioExtend =0, "", Format$(i_RadioExtend)))

Ini_SetKey(s_outputIni, s_section, "ListWidth", Format$(i_ListWidth))
Ini_SetKey(s_outputIni, s_section, "ListHeight", Format$(i_ListHeight))
Ini_SetKey(s_outputIni, s_section, "ListStyle", Format$(i_ListStyle))
Ini_SetKey(s_outputIni, s_section, "ListExtend", Iif$(i_listExtend =0, "", Format$(i_ListExtend)))

Ini_SetKey(s_outputIni, s_section, "MultiWidth", Format$(i_MultiWidth))
Ini_SetKey(s_outputIni, s_section, "MultiHeight", Format$(i_MultiHeight))
Ini_SetKey(s_outputIni, s_section, "MultiStyle", Format$(i_MultiStyle))
Ini_SetKey(s_outputIni, s_section, "MultiExtend", Iif$(i_multiExtend =0, "", Format$(i_MultiExtend)))

Ini_SetKey(s_outputIni, s_section, "EditWidth", Format$(i_EditWidth))
Ini_SetKey(s_outputIni, s_section, "EditHeight", Format$(i_EditHeight))
Ini_SetKey(s_outputIni, s_section, "EditStyle", Format$(i_EditStyle))
Ini_SetKey(s_outputIni, s_section, "EditExtend", Iif$(i_editExtend =0, "", Format$(i_EditExtend)))

Ini_SetKey(s_outputIni, s_section, "MemoWidth", Format$(i_MemoWidth))
Ini_SetKey(s_outputIni, s_section, "MemoHeight", Format$(i_MemoHeight))
Ini_SetKey(s_outputIni, s_section, "MemoStyle", Format$(i_MemoStyle))
Ini_SetKey(s_outputIni, s_section, "MemoExtend", Iif$(i_memoExtend =0, "", Format$(i_MemoExtend)))
Ini_SetKey(s_outputIni, s_section, "StatusExtend", Iif$(i_statusExtend =0, "", Format$(i_statusExtend)))
End Function

Function ControlSetID(ByVal i_ctl As Long) As Long
Local i_return As Long
If a_control(i_ctl) ="button" And a_caption(i_ctl) ="OK" Then
i_return =%ID_OK
ElseIf a_control(i_ctl) ="button" And a_caption(i_ctl) ="Cancel" Then
i_return =%ID_CANCEL
Else
i_return =100 +i_ctl
End If
Function =i_return
End Function

Function ControlSetStyle() As Long
Local i_return As Long

Select Case As Const$ a_control(i_ctl)
Case "label"
i_return =i_labelStyle
Case "button"
i_return =i_buttonStyle
Case "check"
i_return =i_checkStyle
Case "radio"
i_return =i_radioStyle
Case "list"
i_return =i_listStyle
Case "multi"
i_return =i_multiStyle
Case "edit"
i_return =i_editStyle
Case "memo"
i_return =i_memoStyle
End Select
Function =i_return
End Function

Function ControlSetExtend(ByVal i_ctl As Long) As Long
Local i_return As Long

Select Case As Const$ a_control(i_ctl)
Case "label"
i_return =i_labelExtend
Case "button"
i_return =i_buttonExtend
Case "check"
i_return =i_checkExtend
Case "radio"
i_return =i_radioExtend
Case "list"
i_return =i_listExtend
Case "multi"
i_return =i_multiExtend
Case "edit"
i_return =i_editExtend
Case "memo"
i_return =i_memoExtend
End Select
Function =i_return
End Function

Function ControlSetWidth() As Long
Local i_return As Long

Select Case As Const$ a_control(i_ctl)
Case "label"
i_return =i_labelWidth
Case "button"
i_return =i_buttonWidth
Case "check"
i_return =i_checkWidth
Case "radio"
i_return =i_radioWidth
Case "list"
i_return =i_listWidth
Case "multi"
i_return =i_multiWidth
Case "edit"
i_return =i_editWidth
Case "memo"
i_return =i_memoWidth
End Select
Function =i_return
End Function

Function ControlHeightSet() As Long
Local i_return As Long

Select Case As Const$ a_control(i_ctl)
Case "label"
i_return =i_labelHeight
Case "button"
i_return =i_buttonHeight
Case "check"
i_return =i_checkHeight
Case "radio"
i_return =i_radioHeight
Case "list"
i_return =i_listHeight
Case "multi"
i_return =i_multiHeight
Case "edit"
i_return =i_editHeight
Case "memo"
i_return =i_memoHeight
End Select
Function =i_return
End Function

Function ControlSaveSettings(ByVal i_ctl As Long) As Long
ini_SetKey(s_outputIni, a_name(i_ctl), "caption", a_caption(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "band", Iif$(i_ctl, Format$(a_band(i_ctl)), ""))
ini_SetKey(s_outputIni, a_name(i_ctl), "control", a_control(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "id", a_id(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "type", a_type(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "mask", a_mask(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "range", a_range(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "value", IIf$(a_control(i_ctl) ="memo", "", a_value(i_ctl)))
ini_SetKey(s_outputIni, a_name(i_ctl), "selection", a_select(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "focus", a_focus(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "align", a_align(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "left", a_left(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "top", a_top(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "width", a_width(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "height", a_height(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "fore", IIf$(((i_ctl =0) Or (A_fore(i_ctl) <>a_fore(0))), a_fore(i_ctl), ""))
ini_SetKey(s_outputIni, a_name(i_ctl), "back", IIf$(((i_ctl =0) Or (A_back(i_ctl) <>a_back(0))), a_back(i_ctl), ""))
ini_SetKey(s_outputIni, a_name(i_ctl), "style", a_style(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "extend", a_extend(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "tip", a_tip(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "help", a_help(i_ctl))
ini_SetKey(s_outputIni, a_name(i_ctl), "misc", a_misc(i_ctl))
Function =1
End Function

Function ControlIsSet(ByVal i_ctl As Long) As Long
Function = IsFalse (Len(a_left(i_ctl)) =0 Or Len(a_top(i_ctl)) =0 Or Len(a_width(i_ctl)) =0 Or Len(a_height(i_ctl)) =0)
End Function

Function ControlSet(ByVal i_ctl As Long) As Long
a_left(i_ctl) =Format$(x)
a_top(i_ctl) =Format$(y)
a_width(i_ctl) =Format$(xx)
a_height(i_ctl) =Format$(yy)
Function =1
End Function

Function DialogThread(ByVal i_thread As DWord) As Long
Local p_thread As ThreadInfo Ptr
Local sSource, sTarget As String
Local h_parent As Dword

p_thread =i_thread
sSource = @p_thread.Source
sTarget = @p_thread.Target
h_parent =@p_thread.Parent
If h_parent =0 Then h_parent =%HWND_DESKTOP
If Len(sSource) =0 Then sSource =s_AppPath
If IsFalse Instr(sSource, Any ":\") Then sSource =s_AppPath +sSource
If IsFalse Instr(sTarget, Any ":\") Then sTarget =s_AppPath +sTarget
' If sTarget = "" Then sTarget = sSource
If IsFile(sSource) Then
If sTarget = "" Then sTarget = sSource
s_inputIni = sSource
s_outputIni = sTarget

s_InputTxt = s_InputIni
s_HelpTxt = s_InputIni
s_outputTxt = s_outputIni
Else
If IsFolder(sSource) And IsFalse InStr(":\", Mid$(sSource,-1)) Then sSource = sSource & "\"
If IsFalse InStr(":\", Mid$(sSource,-1)) Then sSource =sSource +"_"
s_inputIni =sSource +"input.ini"
s_outputIni =sSource +"output.ini"
s_inputTxt =sSource +"input.txt"
s_outputTxt =sSource +"output.txt"
s_helpTxt =sSource +"help.txt"
End If
s_InputIni = PathScan$(FULL, s_inputIni)
' MsgBox s_inputIni
' MsgBox s_outputIni
DialogLoadFromIni()
If s_dlgInput <>"all" Then
DialogGroupControls()
DialogSizeControls()
DialogPositionControls()
End If
i_result =DialogActivate()

If i_result Then
End If
End Function

Function GetSection(ByVal s_body As String, ByVal s_section As String, ByVal s_default As String) As String
Local i_start As Long, i_end As Long, i_len As Long

local s as string
s ="[[" +s_section +"]]" +$CRLF
i_start =InStr(s_body, s)
If i_start Then
i_start =i_start +Len(s)
s =$CrLf +"[["
i_end =InStr(i_start, s_body, s)
If i_end Then
i_len =i_end -i_start
Else
i_len =Len(s_body) -i_start +1
End If
s =Mid$(s_body, i_start, i_len)
Else
s =s_default
End If
Function =s
End Function

Function Ipt_Box(ByVal s_title As String, ByVal s_label As String, ByVal s_value As String) As String
Local i_ctl As Long
Local s_tempIni As ASCIIZ * %MAX_PATH

s_tempIni =s_appPath +"temp_input.ini"
If IsFile(s_tempIni) Then Kill s_tempIni
local s as string
s ="Find in " +a_name(i_ctl)
Ini_SetKey(s_tempIni, s_title, "control", "form")
'Ini_SetKey(s_tempIni, s_title, "align", Format$(h_dlg))
Ini_SetKey(s_tempIni, s_label, "control", "edit")
Ini_SetKey(s_tempIni, s_label, "value", s_value)
Ini_SetKey(s_tempIni, "OK", "control", "button")
Ini_SetKey(s_tempIni, "Cancel", "control", "button")
' oIniForm.RunForm("temp")
s_tempIni =s_appPath +"temp_output.ini"
Function =Ini_GetKey(s_tempIni, "Results", s_label, "")
End Function

Function GetFocusCTL() As Long
Function =ID2Ctl(GetDlgCtrlID(GetFocus()))
End Function

Function GetFocusID() As Long
Function =GetDlgCtrlID(GetFocus())
End Function

Function SayLine() As Long
Local h As DWord

h =GetFocus()
' Function = SayString(Edit_GetLine(h, Edit_LineFromChar(h, -1)))
End Function

Function ControlFind(ByVal s_title As String) As Long
Local z As ASCIIZ * %MAX_PATH
Local i_start As Long, i_end As Long, i_step As Long, i_find As Long
Local i_item, i_itemCount As Long

If InStr(s_title, "Again") Then
' SayString(s_title)
Else
s_find =DialogInput(s_title, "Text in " +a_name(GetFocusCtl()), s_find)
End If
i_id =GetFocusID()
i_ctl =Id2Ctl(i_id)
Select Case As Const$ a_control(i_ctl)
Case "edit", "memo"
local s as string
Control Get Text h_dlg, i_id to S
Control Send h_dlg, i_id, %EM_GETSEL, VarPtr(i_start), VarPtr(i_end)
If InStr(s_title, "Forward") Then
Incr i_start
Else
i_start =-1 *(Len(s) -i_start) -2
End If
i_find =InStr(i_start, LCase$(s), lCase$(s_find))
If i_find Then
i_find =i_find +Len(s_find) -1
Control Send h_dlg, i_id, %EM_SETSEL, i_find, i_find
Control Send h_dlg, i_id, %EM_SCROLLCARET, 0, 0
Else
' SayString("Not found!")
End If
' SayLine()

Case "list", "multi"
Control Send h_dlg, i_id, %LB_GETCOUNT, 0, 0 To i_itemCount
Control Send h_dlg, i_id, %LB_GETCARETIndex, 0, 0 To i_start
If InStr(s_title, "Forward") Then
Incr i_start
i_end =i_itemCount -1
i_step =+1
Else
Decr i_start
i_end =0
i_step =-1
End If

For i_item =i_start to i_end Step i_step
Control Send h_dlg, i_id, %LB_GETTEXT, i_item, VarPtr(z)
If InStr(LCase$(z), lCase$(s_find)) Then Exit For
Next
If i_item >=0 And i_item <i_itemCount Then
Control Send h_dlg, i_id, %LB_SETCARETINDEX, i_item, 0
Else
' SayString("Not found!")
End If
' SayString(z)
End Select
End Function

CallBack Function DialogEvent() As Long
Select Case As Long CbMsg
Case %WM_COMMAND
' Process control notifications
Function =1
Select Case As Long CbCtl
Case %ID_F1
i =CbCtl
i =Id2Ctl(i_id)
HelpActivate(ID2Ctl(i_id))

Case %ID_CONTROL_S
If CbCtlMsg =1 Then
' SayString("save")
Control Post h_dlg, %ID_OK, %BM_CLICK, 0, 0
'DialogSaveToIni()
'Dialog End h_dlg, 1
Function =1
End If
Case %ID_F5
' SayString("F5")
i_id =GetDlgCtrlID(GetFocus())
i =Id2Ctl(i_id)
Control Set Text h_dlg, i_id, A_value(I)
Case %ID_CONTROL_F
ControlFind("Forward Find")
Case %ID_CONTROL_SHIFT_F
ControlFind("Reverse Find")
Case %ID_F3
ControlFind("Forward Again")
Case %ID_SHIFT_F3
ControlFind("Reverse Again")
Case %ID_CONTROL_A
i_id =GetDlgCtrlID(GetFocus())
ListBox Select h_dlg, i_id, 0
Case %ID_CONTROL_SHIFT_A
i_id =GetDlgCtrlID(GetFocus())
ListBox Unselect h_dlg, i_id, 0
End Select
Case %WM_NOTIFY
Case %WM_NCACTIVATE
Static h_focus As Dword
If IsFalse CbWParam Then
' Save control focus
h_focus = GetFocus()

ElseIf h_focus Then
' Restore control focus
SetFocus(h_focus)
h_focus = 0
End If

Case %WM_KEYDOWN
Case %WM_CHAR
Case %WM_KEYUP
Case %WM_SYSCOMMAND
Case %WM_SYSKEYDOWN
Case %WM_SYSCHAR
Case %WM_SYSKEYUP
Case %WM_SIZE
' Dialog has been resized
Control Send CbHndl, %ID_STATUSBAR, CbMsg, CbWParam,CbLParam

Case %WM_INITDIALOG
Control Send h_dlg, %ID_STATUSBAR, %WM_SIZE, CbWParam,CbLParam
'PostMessage(h_lb, %LB_SETSEL, 1, 1)
ListBox Select h_dlg, Val(a_id(i_lb)), Val(a_select(i_lb))
'Control Send h_dlg, i_lb, %WM_KEYDOWN, %VK_DOWN, 0
SetForegroundWindow(h_dlg)
Dim h_Control As DWord
h_Control = GetDlgItem(h_dlg, Val(a_id(2)))
SetFocus(h_control)
SendMessage(h_control, %EM_SETSEL, 0, -1)
Case %WM_DESTROY
End Select
End Function

CallBack Function LabelEvent() As Long
End Function

CallBack Function ButtonEvent() As Long
Select Case As Long CbMsg
Case %WM_COMMAND
Function =1
Select Case As Long CbCtlMsg
Case %BN_SETFOCUS
i_id =CbCtl
i =Id2Ctl(i_id)
local s as string
s =a_tip(i)
Control Send h_dlg, %ID_STATUSBAR, %SB_SETTEXT, 0, StrPtr(s)
Case %BN_CLICKED
Select Case As Const CbCtl
Case %ID_CANCEL
Dialog End CbHndl, 0
Function =1

Case Else
i_id =CbCtl
DialogSaveToIni(iD2Ctl(i_id))
Dialog End h_dlg, 1
Function =1
End Select
End Select
End Select
End Function

CallBack Function CheckEvent() As Long
Select Case As Long CbMsg
Case %WM_COMMAND
Function =1
Select Case As Long CbCtlMsg
Case %BN_SETFOCUS
i_id =CbCtl
i =Id2Ctl(i_id)
local s as string
s =a_tip(i)
Control Send h_dlg, %ID_STATUSBAR, %SB_SETTEXT, 0, StrPtr(s)
End Select
End Select
End Function

CallBack Function RadioEvent() As Long
Select Case As Long CbMsg
Case %WM_COMMAND
Function =1
Select Case As Long CbCtlMsg
Case %BN_SETFOCUS
i_id =CbCtl
i =Id2Ctl(i_id)
local s as string
s =a_tip(i)
Control Send h_dlg, %ID_STATUSBAR, %SB_SETTEXT, 0, StrPtr(s)
End Select
End Select
End Function

CallBack Function ListEvent() As Long
Select Case As Long CbMsg
Case %WM_COMMAND
Function =1
Select Case As Long CbCtlMsg
Case %LBN_SETFOCUS
i_id =CbCtl
i =Id2Ctl(i_id)
local s as string
s =a_tip(i)
Control Send h_dlg, %ID_STATUSBAR, %SB_SETTEXT, 0, StrPtr(s)
End Select
End Select
End Function

CallBack Function MultiEvent() As Long
Select Case As Long CbMsg
Case %WM_COMMAND
Function =1
Select Case As Long CbCtlMsg
Case %LBN_SETFOCUS
i_id =CbCtl
i =Id2Ctl(i_id)
local s as string
s =a_tip(i)
Control Send h_dlg, %ID_STATUSBAR, %SB_SETTEXT, 0, StrPtr(s)
End Select
End Select
End Function

CallBack Function EditEvent() As Long
Select Case As Long CbMsg
Case %WM_COMMAND
Function =1
Select Case As Long CbCtlMsg
Case %EN_SETFOCUS
i_id =CbCtl
i =Id2Ctl(i_id)
local s as string
s =a_tip(i)
Control Send h_dlg, %ID_STATUSBAR, %SB_SETTEXT, 0, StrPtr(s)
End Select
End Select
End Function

CallBack Function MemoEvent() As Long
Select Case As Long CbMsg
Case %WM_COMMAND
Function =1
Select Case As Long CbCtlMsg
Case %EN_SETFOCUS
i_id =CbCtl
i =Id2Ctl(i_id)
local s as string
s =a_tip(i)
Control Send h_dlg, %ID_STATUSBAR, %SB_SETTEXT, 0, StrPtr(s)

End Select
End Select
End Function

Function GetCommandArgCount() AS LONG
LOCAL iIndex AS LONG
FOR iIndex = 1 TO 100
IF COMMAND$(iIndex) = "" THEN EXIT FOR
NEXT
Decr iIndex
Function = iIndex
End Function

Function LogError(sText AS STRING) AS LONG
LOCAL sClipboard AS STRING

IF ISFALSE bDebugMode THEN EXIT FUNCTION
CLIPBOARD GET TEXT TO sClipboard
sClipboard = sClipboard & $CRLF & sText
CLIPBOARD SET TEXT sClipboard
End Function

Function File2String(BYVAL s_file AS ASCIIZ * 256) AS STRING
LOCAL i_size AS LONG, h_file AS LONG, s_return AS STRING

IF LEN(DIR$(s_file, 7)) =0 THEN
s_return =""
ELSE
h_file =FREEFILE
OPEN s_file FOR BINARY AS h_file
i_size =LOF(h_file)
GET$ h_file, i_size, s_return
CLOSE h_file
END IF
Function =s_return
End Function

Function String2File(BYVAL s_text AS STRING, BYVAL s_file AS ASCIIZ * 256) AS LONG
LOCAL i_size AS LONG, h_file AS LONG, i_return AS LONG

IF ISTRUE ISFILE(s_file) THEN KILL s_File
'msgbox format$(len(s_text))
IF LEN(s_text) =0 THEN
'If IsFalse Then
i_return =0
ELSE
h_file =FREEFILE
OPEN s_file FOR BINARY AS h_file
i_size =LEN(s_text)
PUT$ h_file, s_text
CLOSE h_file
i_return =1
END IF
Function =i_return
End Function

Function PrintLine(Z AS STRING) AS LONG
' returns TRUE (non-zero) on success
LOCAL hStdOut AS LONG, nCharsWritten AS LONG
LOCAL w AS STRING
STATIC CSInitialized AS LONG, CS AS CRITICAL_SECTION
IF ISFALSE CSInitialized THEN
InitializeCriticalSection CS
CSInitialized  =  1
END IF
EntercriticalSection Cs
hStdOut      = GetStdHandle (%STD_OUTPUT_HANDLE)
IF hSTdOut   = -1& OR hStdOut = 0&   THEN     ' invalid handle value, coded in line to avoid
' casting differences in Win32API.INC
' test for %NULL added 8.26.04 for Win/XP
AllocConsole
hStdOut  = GetStdHandle (%STD_OUTPUT_HANDLE)
END IF
LeaveCriticalSection CS
w = Z & $CRLF
Function = WriteFile(hStdOut, BYVAL STRPTR(W), LEN(W),  nCharsWritten, BYVAL %NULL)
End Function

Function StringPlural(sText AS STRING, iCount AS LONG) AS STRING
LOCAL sReturn AS STRING

sReturn = sText
IF iCount <> 1 THEN sReturn = sReturn & "s"
Function = sReturn
End Function

Function GetWidth(iNum AS LONG) AS LONG
LOCAL iResult, iLoop, iPower AS LONG

iLoop = 1
WHILE iLoop > 0
iResult = iNum \ (10^iPower)
IF (iResult = 0) OR (iLoop = 100) THEN
iLoop = -1 * iLoop
ELSE
iPower = iPower + 1
END IF
WEND

IF iLoop = -100 THEN
' DialogShow("reached 100 for width", "")
GetWidth = 5
ELSE
GetWidth = iPower
END IF
End Function

Function DialogInput(sTitle AS STRING, sMessage AS STRING, sValue AS STRING) AS STRING
Function = INPUTBOX$(sMessage, sTitle, sValue)
End Function

Function DialogShow(ByVal sTitle AS STRING, ByVal sMessage AS STRING) AS LONG
' show a standard message box

DIM iFlags AS LONG

DialogShow = %True
iFlags = %MB_ICONINFORMATION OR %MB_SYSTEMMODAL
IF sTitle = "" THEN sTitle = "Show"
MSGBOX sMessage, iFlags, sTitle
End Function

Function StringQuote(BYVAL s$) AS STRING
Function = CHR$(34) & s$ & CHR$(34)
End Function

Function DialogConfirm(sTitle AS STRING, sMessage AS STRING, sDefault AS STRING) AS STRING
' Get choice from a standard Yes, No, or Cancel message box

DIM iFlags AS LONG, iChoice AS LONG

DialogConfirm = ""
iFlags = %MB_YESNOCANCEL
iFlags = iFlags OR %MB_ICONQUESTION     ' 32 query icon
iFlags = iFlags OR %MB_SYSTEMMODAL ' 4096   System modal
IF sTitle = "" THEN sTitle = "Confirm"
IF sDefault = "N" THEN iFlags = iFlags OR %MB_DEFBUTTON2
iChoice = MSGBOX(sMessage, iFlags, sTitle)
IF iChoice = %IDCANCEL THEN EXIT FUNCTION

IF iChoice = %IDYES THEN
DialogConfirm = "Y"
ELSE
DialogConfirm = "N"
END IF
End Function

Function Say(sText AS STRING) AS LONG
DIM sExe AS STRING
sExe = exe.path$ & "SayLine.exe"
SHELL(StringQuote(sExe) & sText, 0)
End Function

Function save_File2String(ByVal s_file As Asciiz * %MAX_PATH) As String
Local i_size As Long, h_file As Long, s_return As String

If Len(Dir$(s_file, 7)) =0 Then
s_return =""
Else
h_file =FreeFile
Open s_file For Binary As h_file
i_size =Lof(h_file)
Get$ h_file, i_size, s_return
Close h_file
End If
Function =s_return
End Function

Function StrToUnicode(ByVal x As String) As String
Local Buffer As String
Buffer = Space$(Len(x) * 2)
'
MultiByteToWideChar %CP_ACP, _
%NULL, _
ByVal StrPtr(x),_
Len(x),_
ByVal StrPtr(Buffer),_
Len(Buffer)
Function = Buffer
End Function

Function save_String2File(ByVal s_text As String, ByVal s_file As Asciiz * %MAX_PATH) As Long
Local i_size As Long, h_file As Long, i_return As Long

If Len(s_text) =0 Then
i_return =0
Else
h_file =FreeFile
Open s_file For Binary As h_file
i_size =Len(s_text)
Put$ h_file, s_text
Close h_file
i_return =1
End If
Function =i_return
End Function

Function Append2File(ByVal s_text As String, ByVal s_file As Asciiz * %MAX_PATH) As Long
If IsFile(s_file) Then
s_text =File2String(s_file) +$CRLF +s_text
Kill s_file
End If
Function = String2File(s_text, s_file)
End Function

CLASS IniForm $CIniFormGuid AS COM
INTERFACE IIniForm $IIniFormGuid
INHERIT Dual

Method RunForm <10> Alias "RunForm" (ByVal sSource As String, ByVal sTarget As String) As Long
Local i, i_threadHandle as Long
Local ti as ThreadInfo

sSource = ACode$(sSource)
sTarget = ACode$(sTarget)
ti.Source =sSource
ti.Target =sTarget

' Thread Create DialogThread(VarPtr(ti)) To i_threadHandle
DialogThread(VarPtr(ti))
Method = IsFile(s_outputIni)
End Method

Method ShowResults <15> Alias "ShowResults" () As Long
Local iReturn as Long
Local sBody, sTitle as String

sTitle = "Show"
sBody = FileToString(s_outputIni)
DialogShow(sTitle, sBody)
End Method

Method GetResult <20> Alias "GetResult" (ByVal sKey As String) As String
Local sReturn as String

sKey = ACode$(sKey)
sReturn = ini_GetKey(s_outputIni, "Results", sKey, "")
If sReturn = "" Then sReturn = GetSection(FileToString(s_outputTxt), sKey, "")
Method = UCode$(sReturn)
End Method

End Interface
End Class

Function PBMain()
Local iCount, iReturn as long
Local sSource, sTarget as String
Local ti as ThreadInfo

iCount = GetCommandArgCount()
s_appPath =PathName$(PATH, EXE.Full$)
' MsgBox s_appPath
If iCount >= 1 Then sSource = Command$(1)
If iCount >= 2 Then sTarget = Command$(2)
oIniForm = Class "IniForm"
sSource = UCode$(sSource)
sTarget = UCode$(sTarget)
Object Call oIniForm.RunForm(sSource, sTarget) to iReturn
Function = iReturn
'Profile s_appPath +"dialog.pro"
End Function

