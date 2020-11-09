VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1425
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   1425
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line2 
      X1              =   0
      X2              =   1440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblHide 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide / Show"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblCryptography_Click()

End Sub

Private Sub lblHide_Click()
    If GetAttr(file) = 2 And vbHidden Then
        lblHide.Caption = "Hide"
        SetAttr file, vbNormal
    Else
        lblHide.Caption = "Show"
        SetAttr file, vbHidden
    End If
    Unload Me
End Sub

Private Sub lblSteganography_Click()
    frmOption2.Show
    Unload Me
End Sub
