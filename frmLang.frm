VERSION 5.00
Begin VB.Form frmLang 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video Grabber - [ Language ]"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLang.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   390
      Left            =   3165
      TabIndex        =   2
      Top             =   615
      Width           =   1470
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmLang.frx":030A
      Left            =   1260
      List            =   "frmLang.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   3405
   End
   Begin VB.Label Label1 
      Caption         =   "Language :"
      Height          =   315
      Left            =   195
      TabIndex        =   0
      Top             =   195
      Width           =   1200
   End
End
Attribute VB_Name = "frmLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    saveLangData frmLang.Combo1
    MsgBox langVar(0).strFull(33), vbOKOnly + vbInformation + vbApplicationModal, "Video Grabber"
    Unload Me
End Sub

Private Sub Form_Load()
    InitCommonControlsVB
    loadLangList frmLang.Combo1
End Sub
