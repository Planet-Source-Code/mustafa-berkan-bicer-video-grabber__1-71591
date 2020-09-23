VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "youtube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function downloadYoutubeVideo(fUrl As String)

'http://www.youtube.com/watch?v=xJCn6SCRh6g


    Dim urlUnparsed As String
    Dim urlParsed As String
    
  '  On Local Error GoTo son
    
    urlUnparsed = grabDirectVideoURL(fUrl)
    urlParsed = parseDirectVideoURL(urlUnparsed)
    
    downloadFile urlParsed
'son:
 '   MsgBox xmlDoc.langVar(0).strFull(18), vbOKOnly + vbExclamation + vbApplicationModal, "Video Grabber"
End Function

Function grabDirectVideoURL(fUrl As String)
    Dim ytVideo As New XMLHTTPRequest
    
    Dim url_sta, url_sto, urlp
    Dim fbBuffer
    
    Dim checkWhichTypeOfYoutube
    Dim tmp0
    
    checkWhichTypeOfYoutube = InStr(1, fUrl, "watch?")
    If (checkWhichTypeOfYoutube <= 0) Then
        checkWhichTypeOfYoutube = InStr(1, fUrl, "v/")
        If checkWhichTypeOfYoutube <> 0 Then
            tmp0 = Mid(fUrl, checkWhichTypeOfYoutube + 2, Len(fUrl) - checkWhichTypeOfYoutube - 1)
            fUrl = "http://www.youtube.com/watch?v=" & tmp0
        End If
    End If
    
    ytVideo.open "GET", fUrl, False
    ytVideo.send
    fbBuffer = ytVideo.responseText
    
    url_sta = InStr(1, fbBuffer, "watch_fullscreen")
    url_sto = InStr(url_sta + 17, fbBuffer, ";")

    urlp = Mid(fbBuffer, url_sta + 21, url_sto - url_sta - 22)
    
    grabDirectVideoURL = urlp
    
End Function

Function parseDirectVideoURL(fUrl As String)
        
    showState xmlDoc.langVar(0).strFull(23)
    parseDirectVideoURL = "http://www.youtube.com/get_video?" & fUrl

End Function

Function downloadFile(fParsedFileUrl As String)
    Dim Size As Long, Remaining As Long, FFile As Integer, Chunk() As Byte
    Dim FileName As String, NowSize As Long, Yuzde As Integer
    
    FileName = Right(fParsedFileUrl, Len(fParsedFileUrl) - InStrRev(fParsedFileUrl, "="))
    
    FileName = Replace(FileName, "!", "_")
    FileName = Replace(FileName, "/", "_")
    FileName = Replace(FileName, "\", "_")
    FileName = Replace(FileName, ";", "_")
    FileName = Replace(FileName, ":", "_")
    FileName = Replace(FileName, ",", "_")
    FileName = Replace(FileName, " ", "_")
    FileName = Replace(FileName, "$", "_")
    FileName = Replace(FileName, "+", "_")
    FileName = Replace(FileName, "&", "_")
    FileName = Replace(FileName, "[", "_")
    FileName = Replace(FileName, "]", "_")
    FileName = Replace(FileName, "(", "_")
    FileName = Replace(FileName, ")", "_")
    FileName = Replace(FileName, "{", "_")
    FileName = Replace(FileName, "}", "_")
    FileName = Replace(FileName, "?", "_")
    FileName = Replace(FileName, "*", "_")
    FileName = Replace(FileName, "-", "_")
    FileName = Replace(FileName, "|", "_")
    FileName = Replace(FileName, "<", "_")
    FileName = Replace(FileName, ">", "_")
    FileName = Replace(FileName, ".", "_")
    FileName = Replace(FileName, "æ", "_")
    FileName = Replace(FileName, "ß", "_")
    FileName = Replace(FileName, "€", "_")
    FileName = Replace(FileName, "@", "_")
    FileName = Replace(FileName, "¨", "_")
    FileName = Replace(FileName, "~", "_")
    FileName = Replace(FileName, "´", "_")
    FileName = Replace(FileName, "`", "_")
    FileName = Replace(FileName, "=", "_")
    FileName = Replace(FileName, "%", "_")
    FileName = Replace(FileName, "½", "_")
    FileName = Replace(FileName, "#", "_")
    FileName = Replace(FileName, "£", "_")
    FileName = Replace(FileName, "'", "_")
    FileName = Replace(FileName, """", "_")
    FileName = Replace(FileName, "é", "_")
    
    FileName = LCase(FileName) & ".flv"
    
    Form1.Inet2.Cancel
    
    Form1.Inet2.Execute Trim(fParsedFileUrl), "GET"
    
    Do While Form1.Inet2.StillExecuting
        DoEvents
    Loop
    
    Form1.ProgressBar1.Max = 100
    
    On Local Error GoTo son
    Size = CLng(Form1.Inet2.GetHeader("Content-Length"))
    Remaining = Size
    NowSize = 0
    Form1.Label8.Caption = Size & " Byte"
    
    FFile = FreeFile
    On Local Error Resume Next
    MkDir App.Path + "\downloaded"
    
    Form1.Label9.Caption = App.Path + "\downloaded\" + FileName
    Form1.Command1.Enabled = False
    Form1.Command3.Enabled = False
    Form1.Check1.Enabled = False
    Form1.Text1.Enabled = False
    Form1.Text2.Enabled = False
    Form1.Text3.Enabled = False
    
    showState xmlDoc.langVar(0).strFull(25)
    
    Open App.Path + "\downloaded\" + FileName For Binary Access Write As #FFile
    Do Until Remaining = 0
        If Remaining > 1024 Then
            Chunk = Form1.Inet2.GetChunk(1024, icByteArray)
            Remaining = Remaining - 1024
        Else
            Chunk = Form1.Inet2.GetChunk(Remaining, icByteArray)
            Remaining = 0
        End If
        NowSize = Size - Remaining
        Yuzde = CInt((100 / Size) * NowSize)
        Form1.Label7.Caption = "% " & Yuzde
        Form1.Caption = "Video Grabber | Grabbing : % " & Yuzde
        Form1.Label10.Caption = NowSize & " Byte"
        Form1.Label12.Caption = Remaining & " Byte"
        Form1.ProgressBar1.Value = Yuzde
        Put #FFile, , Chunk
    Loop
    Close #FFile
    
    Form1.Command1.Enabled = True
    Form1.Command3.Enabled = True
    Form1.Check1.Enabled = True
    Form1.Text1.Enabled = True
    Form1.Text2.Enabled = True
    Form1.Text3.Enabled = True
    
    showState xmlDoc.langVar(0).strFull(26)
    
    Form1.Caption = "Video Grabber"
    
    Exit Function
son:
    
    Form1.Command1.Enabled = True
    Form1.Command3.Enabled = True
    Form1.Check1.Enabled = True
    Form1.Text1.Enabled = True
    Form1.Text2.Enabled = True
    Form1.Text3.Enabled = True
    
    showState "N/A"
    
    Form1.Caption = "Video Grabber"
End Function

