VERSION 5.00
Begin VB.Form frmURLBlocker 
   Caption         =   "Mizano - Block URL "
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBlock 
      Caption         =   "Block"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtURL 
      Height          =   495
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Sites to block: "
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmURLBlocker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBlock_Click()
    Open "C:\WINDOWS\system32\drivers\etc\hosts" For Append As #1
    Print #1, txtURL.Text
    Close #1
    MsgBox "URL Blocked successfully"
    Unload Me
End Sub
