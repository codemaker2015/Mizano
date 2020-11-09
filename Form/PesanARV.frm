VERSION 5.00
Begin VB.Form PesanARV 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   0
      Picture         =   "PesanARV.frx":0000
      ScaleHeight     =   2685
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   0
      Width           =   5745
      Begin Mizano.Abutton Abutton1 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   2040
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         BorderColor     =   -2147483627
         BorderColorPressed=   -2147483628
         BorderColorHover=   -2147483627
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Timer Timer3 
         Interval        =   20000
         Left            =   5160
         Top             =   3120
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   4680
         Top             =   3120
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   4200
         Top             =   3120
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "The datbase is out of date , download Mizano signature database                      "
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Image Picture2 
         Height          =   810
         Left            =   240
         Picture         =   "PesanARV.frx":657E
         Top             =   240
         Width           =   885
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   480
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   5295
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "PesanARV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare the API Calls for Makin Transparent Window


'******************************************************************************************
'******************************************************************************************
'************************************HELLO PROGRAMMERS*************************************
'******************************************************************************************
'******************************************************************************************
'******************************************************************************************


 '   Hello My Name is Shoaib Mohammed and Iam Doin my I BE on Electrical & Electronics
'Engineering.This Small Program Show how to create a High Class Notification Box on The
'task bar as some programs do (example Norton Antivirus). It was Norton which inspired me
'create this. This is not fully error free and needs a lot of maintainance to produce
'a error free message.

'    You have to constantly monitor the message showed. U cannot show a message until
'a previous one  closes completely. other than that there would be no problem , i think.
'i created this program in a hurry ( i am getting late for the date....just kidding...have
'to hang out with friends and prepare for tommororws exam)and so i was no able to add any
'detailed comments.
    
'    The Best Part of this is that it does no use any api call ( Except to make the
'form transparent). It uses Pure VB and Mathematical calculation (easy to undr stand).


    'I think This woudld be much useful and would provide a Better Interface for ur'
'program.

 '   This is just a base class and u are free to change it any way as you like to
'your own style. U have the full permission of using this code or modifying it or using
'it in any commercial package(but please send ur comments that what keeps programmers
'like me on the GO.)


 '   PLEASE SEND IN UR COMMENTS OR BUGS TO

  '      shoaib_134@ rediffmail.com

'*********************************** BYE *******************************************

Const n = vbNewLine
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' Flag to Prevent further Loading of form when it is already on memory
Private Loaded As Boolean


' Transparency Constants
Const LWA_COLORKEY = &H3
Const LWA_ALPHA = &H3
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOOWNERZORDER = &H200
Dim ret As Long

'Form Unloading Area
Private Sub Command1_Click()

End Sub

Private Sub Abutton1_Click()
'set that the form was  loaded once
Loaded = True
'enable the Timer to hide the form
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
'Get the Picture
'Image2.Picture = LoadPicture(App.Path & "\error.gif")

'Set The Position of the form
Me.Top = Screen.Height
Me.Left = Screen.Width - Me.Width



'Check if The form had been loaded , if loaded then unload it
If Loaded Then Unload Me: Loaded = False Else Loaded = True
End Sub

'This sub calculates the total height necessary for the form
'This sub resizes the form according to the message. This should me called before
'-the form is shown
Public Sub Resize()
Me.Height = Me.Height + Label2.Height
Command1.Top = Me.Height - 500 '+ Label2.Height - 500
End Sub


Private Sub Form_Terminate()
Unload Me 'unload the form
End Sub

'Slide the form into view
Private Sub Timer1_Timer()
'Check if the form has reached its maximum height & if Yes then stop timer
If Me.Top <= Screen.Height - (Me.Height + 250) Then Timer1.Enabled = False

'else Move it position to 100 pix top
Me.Top = Me.Top - 100
End Sub
 
 'Hide the form by scroll method ( same as above)
 Public Sub Delete()
 Command1_Click
 End Sub
 
 'Scroll out timer
Private Sub Timer2_Timer()
On Error Resume Next ' Will raise an error if the form is unloaded unexpectedly
'so keep an error trapper

'Check if the form has reached its minimum height & if Yes then stop timer
If Me.Top >= 11000 Then
Timer2.Enabled = False
'Unload the form
Unload Me
End If

'else Move it position to 100 pix down
Me.Top = Me.Top + 100
End Sub


'Time out Timer
'Automatically close the box after 20 secs.
Private Sub Timer3_Timer()

'Enable Scroll out timer
Timer2.Enabled = True

'Disable this timer
Timer3.Enabled = False

End Sub


