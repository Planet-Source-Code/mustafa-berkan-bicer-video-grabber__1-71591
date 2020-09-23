VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Video Grabber"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   " Facebook Information "
      Height          =   1035
      Left            =   120
      TabIndex        =   8
      Top             =   870
      Width           =   8040
      Begin VB.CheckBox Check1 
         Caption         =   "Save My Data"
         Height          =   240
         Left            =   4860
         TabIndex        =   23
         Top             =   675
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4860
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   300
         Width           =   3060
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Left            =   735
         MaxLength       =   255
         TabIndex        =   10
         Top             =   300
         Width           =   3060
      End
      Begin VB.Label Label15 
         Caption         =   "Why I have to supply these?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   735
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   705
         Width           =   3120
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   240
         Left            =   3960
         TabIndex        =   11
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   240
         Left            =   195
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   405
      Top             =   4935
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6180
      TabIndex        =   7
      Top             =   4320
      Width           =   1965
   End
   Begin VB.Frame Frame1 
      Caption         =   " File Information "
      Height          =   2355
      Left            =   120
      TabIndex        =   3
      Top             =   1935
      Width           =   8040
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   300
         Left            =   2895
         TabIndex        =   13
         Top             =   1965
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   0
         Min             =   1e-4
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remain :"
         Height          =   240
         Left            =   195
         TabIndex        =   20
         Top             =   1425
         Width           =   1380
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   1620
         TabIndex        =   19
         Top             =   1425
         Width           =   3075
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current :"
         Height          =   240
         Left            =   195
         TabIndex        =   18
         Top             =   1695
         Width           =   1380
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   1620
         TabIndex        =   17
         Top             =   1695
         Width           =   3075
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   1635
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   16
         Top             =   300
         Width           =   6300
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   1620
         TabIndex        =   15
         Top             =   1170
         Width           =   3075
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   1635
         TabIndex        =   14
         Top             =   2010
         Width           =   1245
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Downloaded :"
         Height          =   240
         Left            =   195
         TabIndex        =   6
         Top             =   1995
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "File Size :"
         Height          =   240
         Left            =   195
         TabIndex        =   5
         Top             =   1170
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Local File :"
         Height          =   240
         Left            =   195
         TabIndex        =   4
         Top             =   300
         Width           =   1380
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grab!"
      Height          =   375
      Left            =   7170
      TabIndex        =   2
      Top             =   360
      Width           =   990
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   105
      MaxLength       =   255
      TabIndex        =   1
      Top             =   375
      Width           =   7035
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   1320
      Top             =   4950
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status :"
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   4365
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      Height          =   255
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   2910
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuLang 
         Caption         =   "Language"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuWeb 
         Caption         =   "Web Site"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        WriteConfFile
    End If
End Sub

Private Sub Command1_Click()
    Dim checkSiteName As String

        checkSiteName = checkSite(Text1.Text)
        If checkSiteName = "youtube" Then
            Dim fbYoutubePage As New youtube
            fbYoutubePage.downloadYoutubeVideo Text1.Text
        ElseIf checkSiteName = "facebook" Then
            If (Text2.Text <> "" And Text3.Text <> "") Then
                    Dim fbFacebookPage As New facebook
                    
                    Text1.Text = Text1.Text + "&"
                    fbFacebookPage.downloadFacebookVideo Text1.Text
                Else
                    MsgBox xmlDoc.langVar(0).strFull(16), vbExclamation + vbOKOnly + vbSystemModal, "Video Grabber"
                    Text2.SetFocus
            End If
        Else
            '
        End If
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    InitCommonControlsVB
    ReadConfFile
    loadLangXMLDoc
End Sub

Private Sub Label14_DblClick()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteConfFile
    End
End Sub

Private Sub Label15_DblClick()
    MsgBox xmlDoc.langVar(0).strFull(6), vbOKOnly + vbInformation + vbApplicationModal, "Video Grabber"
End Sub

Private Sub Label9_DblClick()
    If Label9.Caption <> "" Then
        Shell "explorer.exe """ & Label9.Caption & """"
    End If
End Sub

Private Sub mnuAraclar_Click()

End Sub

Private Sub mnuAbout_Click()
    MsgBox "© 2008, Mustafa Berkan BICER", vbOKOnly + vbInformation, "Author"
End Sub

Private Sub mnuLang_Click()
    frmLang.Show 1
End Sub

Private Sub mnuWeb_Click()
    Shell "explorer.exe ""http://www.mustafaberkanbicer.com.tr"""
End Sub

Private Sub Text1_DblClick()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_DblClick()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text3_DblClick()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
End Sub
