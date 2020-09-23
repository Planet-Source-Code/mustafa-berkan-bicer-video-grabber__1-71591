Attribute VB_Name = "xmlDoc"
Dim langFileContent As New MSXML2.DOMDocument

Public Type langInfo
    langName As String
    langAuthorName As String
    langAuthorMail As String
    strID(33) As String
    strShort(33) As String
    strFull(33) As String
End Type

Public langVar() As langInfo

Function loadLangXMLDoc()

    Dim confFile As String
    Dim vglang As String
    
    confFile = App.Path + "\config\lang.cfg"
    
    Open confFile For Input As #1
        Input #1, vglang
    Close #1
    
    vglang = grabData(vglang)
    
    langFileContent.async = False
    langFileContent.Load App.Path + "\lang\" + vglang + ".xml"
    
    loadSections
    
    applyToForm
End Function

Public Function loadSections()
    Dim objElem As MSXML2.IXMLDOMElement
    Dim objNode As MSXML2.IXMLDOMNodeList
    
    ReDim langVar(1)
        
    Dim langRow As Integer
    
    Set objElem = langFileContent.selectSingleNode("vGrabberLang/langProperties/langName")
    langVar(0).langName = objElem.Text
    Set objElem = langFileContent.selectSingleNode("vGrabberLang/langProperties/langAuthorName")
    langVar(0).langAuthorName = objElem.Text
    Set objElem = langFileContent.selectSingleNode("vGrabberLang/langProperties/langAuthorMail")
    langVar(0).langAuthorMail = objElem.Text
    
    Set objNode = langFileContent.selectNodes("vGrabberLang/langStrings/langString")

    For langRow = 0 To objNode.length - 1
        langVar(0).strID(langRow) = objNode.Item(langRow).childNodes(0).Text
        langVar(0).strShort(langRow) = objNode.Item(langRow).childNodes(1).Text
        langVar(0).strFull(langRow) = objNode.Item(langRow).childNodes(2).Text
    Next langRow

End Function

Function applyToForm()

    Form1.Label1.Caption = langVar(0).strFull(0)
    Form1.Command1.Caption = langVar(0).strFull(1)
    Form1.Frame3.Caption = " " & langVar(0).strFull(2) & " "
    Form1.Label5.Caption = langVar(0).strFull(3)
    Form1.Label6.Caption = langVar(0).strFull(4)
    Form1.Label15.Caption = langVar(0).strFull(5)
    Form1.Check1.Caption = langVar(0).strFull(7)
    Form1.Frame1.Caption = " " & langVar(0).strFull(8) & " "
    Form1.Label2.Caption = langVar(0).strFull(9)
    Form1.Label3.Caption = langVar(0).strFull(10)
    Form1.Label13.Caption = langVar(0).strFull(11)
    Form1.Label11.Caption = langVar(0).strFull(12)
    Form1.Label4.Caption = langVar(0).strFull(13)
    Form1.Label16.Caption = langVar(0).strFull(14)
    Form1.Command3.Caption = langVar(0).strFull(15)
    
    Form1.mnuTools.Caption = langVar(0).strFull(27)
    Form1.mnuLang.Caption = langVar(0).strFull(28)
    Form1.mnuHelp.Caption = langVar(0).strFull(29)
    Form1.mnuWeb.Caption = langVar(0).strFull(30)
    Form1.mnuAbout.Caption = langVar(0).strFull(31)
    
    frmLang.Label1.Caption = langVar(0).strFull(28)
    frmLang.Caption = "Video Grabber - [ " & langVar(0).strFull(28) & " ]"
    frmLang.Command1.Caption = langVar(0).strFull(32)
    
End Function
