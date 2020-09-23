VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "facebook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function downloadFacebookVideo(fUrl As String)
    Dim urlFBParsed As String
    Dim urlVDParsed As String
    Dim urlUnparsed As String
    
  '  On Local Error GoTo son
    
    urlFBParsed = parseFacebookUrl(fUrl)
    urlUnparsed = getFacebookVideoUPLink(urlFBParsed)
    urlVDParsed = parseVideoUrl(urlUnparsed)
    
    downloadFile urlVDParsed
'son:
 '   MsgBox xmlDoc.langVar(0).strFull(18), vbOKOnly + vbExclamation + vbApplicationModal, "Video Grabber"
End Function

Function parseFacebookUrl(fUrl As String)
    Dim videoID As String
    Dim videoIDPos_sta As String
    Dim videoIDPos_sto As String
    
    showState xmlDoc.langVar(0).strFull(19)
    
    videoIDPos_sta = InStr(1, fUrl, "v=")
    videoIDPos_sto = InStr(videoIDPos_sta + 1, fUrl, "&")
    
    videoID = Mid(fUrl, videoIDPos_sta + 2, videoIDPos_sto - videoIDPos_sta - 2)
    
    parseFacebookUrl = "http://www.facebook.com/video/video.php?v=" & videoID
End Function

Function getFacebookVideoUPLink(fUrl As String)

    Dim fbBuffer
    Dim oid_sta, oid_sto, oid
    Dim url_sta, url_sto, urlp
    
    Dim fbLoginReq As New XMLHTTPRequest
    
    showState xmlDoc.langVar(0).strFull(20)
    
    fbLoginReq.open "POST", "https://login.facebook.com/login.php?", False
    fbLoginReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    fbLoginReq.send "email=" & URLEncode(Form1.Text2.Text) & "&pass=" & URLEncode(Form1.Text3.Text)
    
    showState xmlDoc.langVar(0).strFull(21)
    
    fbLoginReq.open "GET", fUrl, False
    fbLoginReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    fbLoginReq.setRequestHeader "Referer", fUrl
    fbLoginReq.send
    fbBuffer = fbLoginReq.responseText
    
    
    oid_sta = InStr(1, fbBuffer, "oid")
    oid_sto = InStr(oid_sta + 3, fbBuffer, ")")

    oid = Mid(fbBuffer, oid_sta + 4, oid_sto - oid_sta - 5)
    
    If (Len(oid) > 15) Then
        oid = ""
    End If
    
    showState xmlDoc.langVar(0).strFull(22)
    
    fbLoginReq.open "GET", fUrl & "&oid=" & oid, False
    fbLoginReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    fbLoginReq.setRequestHeader "Referer", fUrl & "&oid=" & oid
    fbLoginReq.send
    fbBuffer = fbLoginReq.responseText
    
    url_sta = InStr(1, fbBuffer, "video_src")
    url_sto = InStr(url_sta + 9, fbBuffer, ")")

    urlp = Mid(fbBuffer, url_sta + 13, url_sto - url_sta - 14)
    
    showState xmlDoc.langVar(0).strFull(23)
    
    getFacebookVideoUPLink = urlp

End Function
Function parseVideoUrl(urlUnparsed As String)
    Dim urlParsed As String
    
    showState xmlDoc.langVar(0).strFull(24)
    
    urlParsed = Replace(urlUnparsed, "%3A", ":")
    urlParsed = Replace(urlParsed, "%2F", "/")
    
    parseVideoUrl = urlParsed
End Function

Function downloadFile(fParsedFileUrl As String)
    Dim Size As Long, Remaining As Long, FFile As Integer, Chunk() As Byte
    Dim FileName As String, NowSize As Long, Yuzde As Integer
    
    FileName = Right(fParsedFileUrl, Len(fParsedFileUrl) - InStrRev(fParsedFileUrl, "/"))
    
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
