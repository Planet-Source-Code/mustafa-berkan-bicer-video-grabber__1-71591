Attribute VB_Name = "parse_funcs"
Option Explicit

Type fileInfo
    id As String
    url As String
    directurl As String
    title As String
End Type

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function


Function checkSite(vUrl As String)
    Dim checkSiteName As Integer
    
    checkSiteName = InStr(1, vUrl, "youtube.com")
    If checkSiteName <= 0 Then
        checkSiteName = InStr(1, vUrl, "facebook.com")
        If checkSiteName <= 0 Then
            MsgBox xmlDoc.langVar(0).strFull(17), vbOKOnly + vbInformation + vbSystemModal, "Video Grabber"
        Else
            checkSite = "facebook"
        End If
    Else
       checkSite = "youtube"
    End If
End Function

Function ReadConfFile()
    Dim confFile As String
    Dim fbUser As String
    Dim fbPass As String
    Dim fbSaveSettings As String
    
    confFile = App.Path + "\config\config.cfg"
    
    Dim freeFileID
    freeFileID = FreeFile
    
    Open confFile For Input As #freeFileID
        Input #freeFileID, fbUser
        Input #freeFileID, fbPass
        Input #freeFileID, fbSaveSettings
    Close #freeFileID
    
    fbUser = URLDecode(grabData(fbUser))
    fbPass = URLDecode(grabData(fbPass))
    fbSaveSettings = grabData(fbSaveSettings)
    
    Form1.Text2.Text = fbUser
    Form1.Text3.Text = fbPass
    If fbSaveSettings = "yes" Then
        Form1.Check1.Value = vbChecked
    Else
        Form1.Check1.Value = vbUnchecked
    End If
    
End Function

Function WriteConfFile()
    Dim confFile As String
    Dim fbUser As String
    Dim fbPass As String
    Dim fbSaveSettings As String
    
    confFile = App.Path + "\config\config.cfg"
    
    
    If Form1.Check1.Value = vbChecked Then
        fbUser = URLEncode(Form1.Text2.Text)
        fbPass = URLEncode(Form1.Text3.Text)
        fbSaveSettings = "yes"
    Else
        fbUser = ""
        fbPass = ""
        fbSaveSettings = "no"
    End If
    
    Dim freeFileID
    freeFileID = FreeFile
    
    Open confFile For Output As #freeFileID
        Print #freeFileID, "fbUser=" & fbUser
        Print #freeFileID, "fbPass=" & fbPass
        Print #freeFileID, "fbSaveSettings=" & fbSaveSettings
    Close #freeFileID
    
End Function
Function loadLangList(objCombo As ComboBox)
    Dim confFile As String
    
    Dim vgLangNames As String
    Dim vgLangFiles As String
    
    Dim vgLangNamesVar() As String
    Dim vgLangFilesVar() As String
    
    Dim langRow
    
    confFile = App.Path + "\config\lang_list.cfg"
    
    Dim freeFileID
    freeFileID = FreeFile
    
    Open confFile For Input As #freeFileID
        Input #freeFileID, vgLangNames
        Input #freeFileID, vgLangFiles
    Close #freeFileID
    
    vgLangNames = grabData(vgLangNames)
    vgLangFiles = grabData(vgLangFiles)
    
    vgLangNamesVar = Split(vgLangNames, ";")
    vgLangFilesVar = Split(vgLangFiles, ";")
    
    For langRow = 0 To UBound(vgLangNamesVar)
        objCombo.AddItem Trim(vgLangNamesVar(langRow)) & "=" & Trim(vgLangFilesVar(langRow))
    Next langRow
    
End Function

Function saveLangData(objCombo As ComboBox)
    Dim confFile As String
    Dim strLangFile
    
    strLangFile = grabData(objCombo.Text)

    confFile = App.Path + "\config\lang.cfg"
    
    Dim freeFileID
    freeFileID = FreeFile
    
    Open confFile For Output As #freeFileID
        Print #freeFileID, "vgLang=" & strLangFile
    Close #freeFileID
End Function

Function grabData(strData As String)
    Dim strStart
    strStart = InStr(1, strData, "=")
    grabData = Mid(strData, strStart + 1, Len(strData) - strStart)
End Function

Public Function URLEncode(strEntrada As String) As String

Dim i As Long
Dim strSalida As String
Dim Temp As String

For i = 1 To Len(strEntrada)
Temp = Mid(strEntrada, i, 1)
'si queremos que convierta TODOS los caracteres, comentamos las lineas
'que tienen un ## al final
If Not Temp Like "[a-z,A-Z,0-9]" Then '##
strSalida = strSalida & "%" & Hex(Asc(Temp))
Else '##
strSalida = strSalida & Temp '##
End If '##
Next i

URLEncode = strSalida

End Function

Public Function URLDecode(strEntrada As String) As String

Dim strCaracter  As String
Dim strSalida As String
Dim i As Long

For i = 1 To Len(strEntrada)
If Mid(strEntrada, i, 1) = "%" Then

strCaracter = Mid(strEntrada, i + 1, 2)

strSalida = strSalida & Chr(Val("&H" & strCaracter))

i = i + 2

Else

strSalida = strSalida & Mid(strEntrada, i, 1)

End If
Next i

URLDecode = strSalida

End Function

Function showState(strState)
    DoEvents
    Form1.Label16.Caption = xmlDoc.langVar(0).strFull(14) & " " & strState
    DoEvents
End Function
