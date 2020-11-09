VERSION 5.00
Begin VB.Form ATV 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ATV Guard - 1.0.3  BETA"
   ClientHeight    =   8970
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9825
   Icon            =   "ATV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ControlPanel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6930
      Left            =   3480
      Picture         =   "ATV.frx":0CCA
      ScaleHeight     =   6930
      ScaleWidth      =   8505
      TabIndex        =   5
      Top             =   360
      Width           =   8505
      Begin VB.Frame fraEmpty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version Information"
         Enabled         =   0   'False
         Height          =   3135
         Index           =   13
         Left            =   5400
         TabIndex        =   6
         Top             =   1560
         Width           =   2895
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   13
            Left            =   240
            Picture         =   "ATV.frx":7248
            Top             =   1920
            Width           =   240
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resident Shield : V.3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Left            =   1200
            TabIndex        =   13
            Top             =   1440
            Width           =   1485
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00C0C0C0&
            X1              =   240
            X2              =   720
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Firewall 1.0.3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   36
            Left            =   720
            TabIndex        =   12
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Process Explorer 1.0.3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   35
            Left            =   720
            TabIndex        =   11
            Top             =   2280
            Width           =   1635
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Power Removal   1.0.3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   34
            Left            =   720
            TabIndex        =   10
            Top             =   1920
            Width           =   1635
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Signature: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   33
            Left            =   1200
            TabIndex        =   9
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Engine V.3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   32
            Left            =   1200
            TabIndex        =   8
            Top             =   720
            Width           =   765
         End
         Begin VB.Label lblEmpty 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Virus Scanner 1.0.3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   31
            Left            =   600
            TabIndex        =   7
            Top             =   360
            Width           =   1410
         End
         Begin VB.Line linSMP 
            BorderColor     =   &H00C0C0C0&
            Index           =   0
            X1              =   240
            X2              =   240
            Y1              =   720
            Y2              =   1560
         End
         Begin VB.Line linSMP 
            BorderColor     =   &H00C0C0C0&
            Index           =   5
            X1              =   240
            X2              =   720
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line linSMP 
            BorderColor     =   &H00C0C0C0&
            Index           =   7
            X1              =   240
            X2              =   720
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   2
            Left            =   120
            Picture         =   "ATV.frx":77D2
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   4
            Left            =   840
            Picture         =   "ATV.frx":7D5C
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   195
            Index           =   6
            Left            =   840
            Picture         =   "ATV.frx":82E6
            Top             =   1440
            Width           =   195
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   5
            Left            =   240
            Picture         =   "ATV.frx":861A
            Top             =   2640
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   7
            Left            =   240
            Picture         =   "ATV.frx":8BA4
            Top             =   2280
            Width           =   240
         End
         Begin VB.Image imgSMP 
            Height          =   240
            Index           =   3
            Left            =   840
            Picture         =   "ATV.frx":912E
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Resident Shield "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "ATV.frx":96B8
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Get Complete Protection Security or update your database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         TabIndex        =   26
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "http://rexsonic-technologie.webs.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         MouseIcon       =   "ATV.frx":980A
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   5880
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   480
         Picture         =   "ATV.frx":995C
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanning Now "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "ATV.frx":9F32
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   360
         Picture         =   "ATV.frx":A084
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Firewall"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "ATV.frx":A5EE
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   855
         Left            =   480
         Picture         =   "ATV.frx":A740
         Top             =   3960
         Width           =   810
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Setting "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "ATV.frx":AD1F
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Image Image4 
         Height          =   810
         Left            =   480
         Picture         =   "ATV.frx":AE71
         Top             =   5160
         Width           =   870
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   3240
         Picture         =   "ATV.frx":B46B
         Top             =   5280
         Width           =   480
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         MouseIcon       =   "ATV.frx":B962
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Image6 
         Height          =   720
         Left            =   3120
         Picture         =   "ATV.frx":BAB4
         Top             =   4080
         Width           =   720
      End
      Begin VB.Image Image9 
         Height          =   720
         Left            =   3120
         Picture         =   "ATV.frx":C0CB
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image Image7 
         Height          =   615
         Left            =   3120
         Picture         =   "ATV.frx":C73C
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "User Status "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         MouseIcon       =   "ATV.frx":CC5E
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Update Now"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         MouseIcon       =   "ATV.frx":CDB0
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Quarantine"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         MouseIcon       =   "ATV.frx":CF02
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         X1              =   3600
         X2              =   9600
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Component From Control Panel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         X1              =   3600
         X2              =   9720
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Description Or Selected Components"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   480
         TabIndex        =   14
         Top             =   6240
         Width           =   3135
      End
   End
   Begin VB.PictureBox ATVGuardAntiTrojan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6930
      Left            =   120
      Picture         =   "ATV.frx":D054
      ScaleHeight     =   6930
      ScaleWidth      =   10785
      TabIndex        =   1
      Top             =   360
      Width           =   10785
      Begin VB.PictureBox sICON 
         Height          =   255
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picTmpIcon 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   37
         ToolTipText     =   "Pengaturan PicTmpIcon HArus seperti in ( Standarnya )"
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Statistics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   28
         Top             =   5280
         Width           =   3015
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Update "
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Virus Database"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "ARV Version "
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "License To"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "Freeware"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   32
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "V.3"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   31
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   30
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   29
            Top             =   360
            Width           =   1455
         End
      End
      Begin ATVGuard.Abutton a3 
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   2520
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         ButtonStyle     =   7
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Control Panel"
         HandPointer     =   -1  'True
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
      Begin ATVGuard.Abutton a2 
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   3120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         ButtonStyle     =   7
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Registry Fixer"
         HandPointer     =   -1  'True
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
      Begin ATVGuard.Abutton a1 
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         ButtonStyle     =   7
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "Scanning"
         HandPointer     =   -1  'True
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
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Simple , Fast  and  Clean"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "ATV Guard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   360
         Picture         =   "ATV.frx":135D2
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   7440
      Width           =   615
   End
   Begin VB.Menu MnFile 
      Caption         =   "File"
      Begin VB.Menu MnExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnCMOP 
      Caption         =   "Database"
      Begin VB.Menu Msthreat 
         Caption         =   "Signature"
      End
   End
   Begin VB.Menu MnCS 
      Caption         =   "Components"
      Begin VB.Menu MnPW 
         Caption         =   "Power Removal"
      End
      Begin VB.Menu MnRX 
         Caption         =   "Registry Fixer"
      End
      Begin VB.Menu MnGuard 
         Caption         =   "Resident Shield"
      End
      Begin VB.Menu MnFirewall 
         Caption         =   "Firewall"
      End
   End
   Begin VB.Menu MnWT 
      Caption         =   "Windows Tool"
      Begin VB.Menu MnPM 
         Caption         =   "Process Manager"
      End
      Begin VB.Menu MnRGD 
         Caption         =   "Regedit"
      End
      Begin VB.Menu MnSRE 
         Caption         =   "MSConfig"
      End
      Begin VB.Menu MnCMD 
         Caption         =   "CMD"
      End
   End
   Begin VB.Menu MnHelp 
      Caption         =   "Help"
      Begin VB.Menu MnWP 
         Caption         =   "Web Page"
      End
      Begin VB.Menu MnTutorial 
         Caption         =   "Readme"
      End
      Begin VB.Menu MnAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MnStart 
         Caption         =   "Start"
      End
   End
End
Attribute VB_Name = "ATV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REG As New cRegistry
Dim Seal As New ClsHuffman
Dim Col As Collection
Dim Ondata As Collection
Public CekSetting As Boolean, cekLoad As Boolean
Dim X As Integer
Private Sub a1_Click()
Me.Scanner.Visible = True
Me.ControlPanel.Visible = False
Me.Signature.Visible = False
Me.Picture3.Visible = False
Me.RegistryFixer.Visible = False
Me.Picture4.Visible = False
End Sub
Public Sub Message()
Shield.Show
End Sub
Sub LoadDrive()
'    On Error Resume Next
'    Dim LDs As Long, Cnt As Long, sDrives As String
'    LDs = GetLogicalDrives
'    For Cnt = 0 To 25
'        If (LDs And 2 ^ Cnt) <> 0 Then
'            Dim Serial As Long, VName As String, FSName As String, ndrvName As String
'            VName = String$(255, Chr$(0))
'            FSName = String$(255, Chr$(0))
'            GetVolumeInformation Chr$(65 + Cnt) & ":\", VName, 255, Serial, 0, 0, FSName, 255
'            VName = Left$(VName, InStr(1, VName, Chr$(0)) - 1)
'            FSName = Left$(FSName, InStr(1, FSName, Chr$(0)) - 1)
'            ndrvName = ""
'            If VName = "" Then
'                Select Case GetTipeDrive(Chr$(65 + Cnt) & ":\")
'                       Case 2: ndrvName = "3½ Floppy (" & Chr$(65 + Cnt) & ":)"
'                       Case 5: ndrvName = "CDROM (" & Chr$(65 + Cnt) & ":)"
'                       Case Else: ndrvName = "Unknown (" & Chr$(65 + Cnt) & ":)"
'                End Select
'                If ndrvName <> "" Then
'                    lstDrive.AddItem Chr$(65 + Cnt) & ":\" & vbTab & ndrvName
'
'                End If
'            Else
'                ndrvName = VName & " (" & Chr$(65 + Cnt) & ":)"
'                lstDrive.AddItem Chr$(65 + Cnt) & ":\" & vbTab & ndrvName
'                'Chr$(65 + Cnt) & ":\", ndrvName)
'            End If
'        End If
'    Next Cnt

    On Error Resume Next
    Dim LDs As Long, Cnt As Long, sDrives As String
    LDs = GetLogicalDrives
    For Cnt = 0 To 25
        If (LDs And 2 ^ Cnt) <> 0 Then
            lstDrive.AddItem Chr(Cnt + 65) & ":\"
        End If
    Next Cnt
    
    Dim i As Integer
    For i = 0 To lstDrive.ListCount - 1
        lstDrive.Selected(i) = True
    Next
End Sub
Private Sub Quarantine()
On Error Resume Next
    
    Dim nama, Exten As String
    Dim i As Long
    Dim strFile As String, strName As String
    
    With ListView1.ListItems
        For i = 1 To .count
            strFile = .Item(i).SubItems(1)
            If .Item(i).Checked Then
                nama = GetFileName(strFile)
                Exten = Right$(strFile, 3)
                SetFileAttributes nama, FILE_ATTRIBUTE_NORMAL
                DoEvents
                TerminateExeName strFile
                If Seal.EncodeFile(strFile, App.path & "\Quarantine\" & nama & "." & Exten & ".atv") = False Then
                    MsgBox "Virus seal infalid !", vbOKOnly, APP_PROGRAM
                End If
                Open (strFile) For Output As #1
                Close (1)
                Kill (strFile)
                MsgBox " Virus Success Move to Quarantine ", vbInformation, "ATV Guard"
                .Remove i
                Exit Sub
            End If
        Next i
    End With
End Sub

Private Sub a2_Click()
Me.Scanner.Visible = False
Me.ControlPanel.Visible = False
Me.Signature.Visible = False
Me.Picture3.Visible = False
Me.RegistryFixer.Visible = True
Me.Picture4.Visible = False
End Sub

Private Sub a3_Click()
Me.Scanner.Visible = False
Me.ControlPanel.Visible = True
Me.Signature.Visible = False
Me.Picture3.Visible = False
Me.RegistryFixer.Visible = False
Me.Picture4.Visible = False
End Sub

Private Sub Abutton3_Click()
  If Fixed.Enabled = False Then
        If MsgBox("Are you sure want to repair registry ?", vbExclamation + vbYesNo, "ATV Guard") = vbYes Then
            LblFixed.Caption = "Please wait take a few moment..."
            'LockControl False
            Fixed.Enabled = True
            PBFix.value = 0
        End If
    Else
        Fixed.Enabled = False
    End If
End Sub

Private Sub Abutton4_Click()
Quarantine
End Sub

Private Sub Abutton5_Click()
Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    On Error Resume Next    'in case not found, and on cd
    Set drvs = fso.Drives
    For Each drv In drvs
        DoEvents
        Kill drv.DriveLetter & ":\autorun.inf"
    Next
    MsgBox "Autorun.inf files from all drives were removed.", vbInformation, "ATV Guard"
    Set fso = Nothing
    Set drv = Nothing
    Set drvs = Nothing
End Sub

Private Sub Abutton6_Click()
CleanSelected
End Sub
Private Sub CleanSelected()
    If Qrtna.FileName = "" Then
        MsgBox "File not found or file not selected.", vbExclamation, "Quarantine"
    Else
        LogFile "Clean from quarantine folder   " & Qrtna.FileName
        DeleteIt (App.path & "\Quarantine\" & "\" & Qrtna.List(Qrtna.ListIndex))
        Qrtna.Refresh
    End If
End Sub
Private Sub Abutton7_Click()
CleanAll
End Sub
Private Sub CleanAll()
    If Qrtna.FileName = "" Then
        MsgBox "File not found or file not selected.", vbExclamation, "Quarantine"
        Exit Sub
    ElseIf Qrtna.FileName <> "" Then
        If MsgBox("Are you sure clean all object?", vbQuestion + vbYesNo, "/Warning") = vbYes Then
            Kill App.path & "\Quarantine\" & "*.atv"
            MsgBox "All object has been cleaned.", vbInformation, "Quarantine"
            Qrtna.Refresh
        End If
    End If
End Sub


Private Sub Abutton8_Click()
Me.Scanner.Visible = False
Me.ControlPanel.Visible = True
Me.Signature.Visible = False
Me.Picture3.Visible = False
Me.RegistryFixer.Visible = False
Me.Picture4.Visible = False
End Sub

Private Sub Abutton9_Click()
 Dim Adr As String
    
    If Qrtna.FileName = "" Then
        MsgBox "File not found pr file not selected.", vbExclamation, "Quarantine"
    Else
        If MsgBox("Are you sure restore this file?", vbQuestion + vbYesNo, "/Warning") = vbYes Then
           Adr = FileParsePath(App.path & "\Quarantine\" & "\" & Qrtna.List(Qrtna.ListIndex), False, False) & FileParsePath(App.path & "\Quarantine\" & "\" & Qrtna.List(Qrtna.ListIndex), True, False)
            If Seal.DecodeFile(App.path & "\Quarantine\" & "\" & Qrtna.List(Qrtna.ListIndex), Adr) = False Then
                Call MsgBox("Virus Seal Invalid !", vbOKOnly, "Error")
                Exit Sub
            End If
            LogFile "Restore from quarantine folder  " & Qrtna.FileName
            DeleteIt (App.path & "\Quarantine\" & "\" & Qrtna.List(Qrtna.ListIndex))
            Qrtna.Refresh
        End If
    End If
End Sub

Private Sub C_Click()
DeleteNow ListView1, 1
End Sub

Private Sub Check1_Click()
Dim F As Integer
If Check1.value = 1 Then
    For F = 1 To ListView1.ListItems.count
        ListView1.ListItems(F).Checked = True
    Next
Else
    For F = 1 To ListView1.ListItems.count
        ListView1.ListItems(F).Checked = False
    Next
End If
End Sub
Private Sub Command2_Click()
Static jum_Vir As Integer
If Len(Text1.Text) > 0 Then
    If Command2.Caption = "Scan" Then
        Command2.Caption = "Stop"
        ListView1.ListItems.Clear
        Scan (Text1.Text)
        Command2.Caption = "Scan"
    Else
        Command2.Caption = "Scan"
    End If
    jum_Vir = ListView1.ListItems.count
    JumlahFile.Caption = "" & jumlah_file & Chr(13) & ""
    JumlahDirectori.Caption = "" & JumDir & Chr(13) & ""
    MsgBox "Scanning Finished !", vbInformation, "ATV Guard"
Else
    MsgBox "Path not found !", vbCritical, "ATV Guard"
End If
jumlah_file = 0
JumDir = 0
End Sub

Private Sub Command4_Click()
'FrmTest.
End Sub

Private Sub Command5_Click()
Dim BFF As String
BFF = BrowseForFolder(Me.hwnd, _
        "Select Path / Directory to be Scanned :")

        If Len(BFF) > 0 Then
            Text1.Text = BFF
            Command2.Enabled = True
        End If

End Sub

Function Del(mana As String)
SetAttr mana, vbNormal
Kill mana
End Function

Private Sub Fixed_Timer()
If PBFix.value >= PBFix.Max Then
        Fixed.Enabled = False
        FixRegistry
        LblFixed.Caption = "Registry have repairing by ATV Guard"
        'LockControl True
        PBFix.value = 0
    Else
        PBFix.value = PBFix.value + 1
    End If
End Sub

Private Sub Form_Activate()
 Qrtna = App.path & "\Quarantine\"
End Sub

Private Sub Form_Load()
 If App.PrevInstance Then
        MsgBox "ATV Guard are ready run in your system", vbCritical
        End
    End If
Call HitDatabase
    cmdTweak(0).Enabled = False
    cekLoad = False
    CekSetting = False
    GetApp
    cekLoad = True
End Sub

Private Sub Label10_Click()
FrmTest.Show
End Sub

Private Sub Label14_Click()
Me.Scanner.Visible = True
Me.ControlPanel.Visible = False
Me.Signature.Visible = False
Me.Picture3.Visible = False
Me.RegistryFixer.Visible = False
Me.Picture4.Visible = False
End Sub

Private Sub Label16_Click()
Firewall.Show
End Sub

Private Sub Label17_Click()
Me.Scanner.Visible = False
Me.ControlPanel.Visible = False
Me.Signature.Visible = False
Me.Picture3.Visible = False
Me.RegistryFixer.Visible = True
Me.Picture4.Visible = False
End Sub

Private Sub Label19_Click()
About.Show
End Sub

Private Sub Label21_Click()
StatusRegister.Show
End Sub

Private Sub Label22_Click()
MsgBox "Visited Http://rexsonic-technologie.webs.com", vbInformation, "ATV Guard"
End Sub

Private Sub Label23_Click()
Me.Scanner.Visible = False
Me.ControlPanel.Visible = False
Me.Signature.Visible = False
Me.Picture3.Visible = False
Me.RegistryFixer.Visible = False
Me.Picture4.Visible = True
End Sub

Private Sub Label51_Click()
ShellExecute Me.hwnd, vbNullString, "http://rexsonic-technologie.webs.com", vbNullString, "C:\", 1
End Sub

Private Sub lblScan_Change()
lblVir.Caption = ListView1.ListItems.count & " virus"
End Sub

Private Sub lstVirus_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim h() As String
h() = Split(lstVirus.SelectedItem.Text, ".")
lblname1 = LCase(Trim(h(1)))
lbltype2 = LCase(Trim(h(0)))
End Sub

Private Sub MnAbout_Click()
About.Show
End Sub

Private Sub MnCMD_Click()
 Shell MyWindowSys & "cmd.exe", 1
End Sub

Private Sub MnExit_Click()
Me.Hide
End Sub

Private Sub MnFirewall_Click()
Firewall.Show
End Sub

Private Sub MnGuard_Click()
FrmTest.Show
End Sub

Private Sub MnPM_Click()
Shell MyWindowSys & "taskmgr.exe", 1
End Sub

Private Sub MnPW_Click()
PowerRemoval.Show
End Sub

Private Sub MnRGD_Click()
 Shell MyWindowDir & "regedit.exe", 1
End Sub

Private Sub MnRX_Click()
Me.Scanner.Visible = False
Me.ControlPanel.Visible = False
Me.Signature.Visible = False
Me.Picture3.Visible = False
Me.RegistryFixer.Visible = True
Me.Picture4.Visible = False
End Sub

Private Sub MnSRE_Click()
  Shell MyWindowSys & "dllcache\msconfig.exe", 1
End Sub

Private Sub MnStart_Click()
ATV.Show
End Sub

Private Sub MnTutorial_Click()
Shell "notepad.exe " & ahpath(App.path) & "readme.txt", 1
End Sub
Function ahpath(mypath As String) As String
If Right(mypath, 1) = "\" Then
   ahpath = mypath
Else
   ahpath = mypath & "\"
End If
End Function
Private Sub MnWP_Click()
ShellExecute Me.hwnd, vbNullString, "http://rexsonic-technologie.webs.com", vbNullString, "C:\", 1
End Sub
Private Sub Msthreat_Click()
Me.Scanner.Visible = False
Me.ControlPanel.Visible = False
Me.Signature.Visible = True
Me.Picture3.Visible = True
Me.RegistryFixer.Visible = False
End Sub


Private Sub chkSystem_Click(Index As Integer)
    On Error Resume Next
    If cekLoad = True Then
        CekSetting = True
        cmdTweak(0).Enabled = True
        cmdTweak(0).Caption = "Apply"
    End If
End Sub

Sub Apply()
    SaveApp
    cmdTweak(0).Enabled = False
    cmdTweak(0).Caption = "No Changes"
    LockWindowUpdate (GetDesktopWindow())
    ForceCacheRefresh
    LockWindowUpdate (0)
End Sub

Sub Clear()
    Dim i As Integer
    On Error Resume Next
    With chkSystem
        For i = 0 To .count
            .Item(i).value = 0
        Next i
    End With
End Sub

Sub Cek()
    Dim i As Integer
    On Error Resume Next
    With chkSystem
        For i = 0 To .count
            .Item(i).value = 1
        Next i
    End With
End Sub

Private Sub cmdTweak_Click(Index As Integer)
    Select Case Index
        Case 0: Apply
        Case 1: Cek
        Case 2: Clear
        Case 3: Unload Me
    End Select
End Sub


Sub LockControl(bLock As Boolean)
    cmdTweak(0).Enabled = False
    cmdTweak(1).Enabled = bLock
    cmdTweak(2).Enabled = bLock
    cmdTweak(3).Enabled = bLock
    Picture1.Enabled = bLock
    Picture6.Enabled = bLock
    Picture7.Enabled = bLock
    Picture8.Enabled = bLock
End Sub

