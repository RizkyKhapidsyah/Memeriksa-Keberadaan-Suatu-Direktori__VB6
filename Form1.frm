VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Keberadaan Suatu Direktori"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Const ATTR_DIRECTORY = 16
  If Dir$("c:\windows", ATTR_DIRECTORY) <> "" Then
     MsgBox "Direktori ada!", vbInformation, "Ada"
  Else
     MsgBox "Direktori tidak ada!", _
            vbCritical, "Tidak Ada"
  End If
End Sub

Private Sub Form_Load()
    Command1.Caption = "Periksa Direktori"
End Sub


