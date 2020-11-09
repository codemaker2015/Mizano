VERSION 5.00
Begin VB.Form frmRegistryFixer 
   BorderStyle     =   0  'None
   Caption         =   "Registry Fixer"
   ClientHeight    =   9945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   Picture         =   "frmRegistryFixer.frx":0000
   ScaleHeight     =   9945
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox RegistryFixer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6570
      Left            =   5400
      Picture         =   "frmRegistryFixer.frx":19E6D
      ScaleHeight     =   6570
      ScaleWidth      =   8505
      TabIndex        =   0
      Top             =   3600
      Width           =   8505
      Begin VB.Timer Fixed 
         Enabled         =   0   'False
         Interval        =   70
         Left            =   7800
         Top             =   240
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3165
         Left            =   240
         ScaleHeight     =   3165
         ScaleWidth      =   3690
         TabIndex        =   30
         Top             =   1080
         Width           =   3690
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show extensions for known file types"
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
            Height          =   240
            Index           =   16
            Left            =   75
            TabIndex        =   43
            Tag             =   "HideFileExt"
            Top             =   2910
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Hidden Folders And Files "
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
            Height          =   240
            Index           =   15
            Left            =   75
            TabIndex        =   42
            Tag             =   "Hidden "
            Top             =   2670
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show File Hidden Operating System "
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
            Height          =   240
            Index           =   14
            Left            =   75
            TabIndex        =   41
            Tag             =   "ShowSuperHidden "
            Top             =   2430
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Properties My Computer"
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
            Height          =   240
            Index           =   13
            Left            =   75
            TabIndex        =   40
            Tag             =   "NoPropertiesMyComputer"
            Top             =   2190
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Hide And Support"
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
            Height          =   240
            Index           =   12
            Left            =   75
            TabIndex        =   39
            Tag             =   "NoSMHelp"
            Top             =   1950
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Registry Editor Tools"
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
            Height          =   240
            Index           =   1
            Left            =   75
            TabIndex        =   38
            Tag             =   "DisableRegistryTools"
            Top             =   270
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Folder Options Menu"
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
            Height          =   240
            Index           =   2
            Left            =   75
            TabIndex        =   37
            Tag             =   "NoFolderOptions"
            Top             =   510
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Menu Find"
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
            Height          =   240
            Index           =   3
            Left            =   75
            TabIndex        =   36
            Tag             =   "NoFind"
            Top             =   750
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Menu Run"
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
            Height          =   240
            Index           =   4
            Left            =   75
            TabIndex        =   35
            Tag             =   "NoRun"
            Top             =   990
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Right-click on Desktop"
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
            Height          =   240
            Index           =   5
            Left            =   75
            TabIndex        =   34
            Tag             =   "NoViewContextMenu"
            Top             =   1230
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Windows Version on Desktop"
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
            Height          =   240
            Index           =   6
            Left            =   75
            TabIndex        =   33
            Tag             =   "PaintDesktopVersion"
            Top             =   1470
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Display Properties"
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
            Height          =   240
            Index           =   7
            Left            =   75
            TabIndex        =   32
            Tag             =   "NoDispCPL"
            Top             =   1710
            Width           =   3465
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Task Manager"
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
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   31
            Tag             =   "DisableTaskMgr"
            Top             =   75
            Width           =   3465
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3465
         Left            =   4560
         ScaleHeight     =   3465
         ScaleWidth      =   3615
         TabIndex        =   15
         Top             =   1080
         Width           =   3615
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show Full Path at Address Bar"
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
            Height          =   240
            Index           =   35
            Left            =   0
            TabIndex        =   29
            Tag             =   "FullPathAddress"
            Top             =   3000
            Width           =   3390
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the File System Button "
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
            Height          =   240
            Index           =   34
            Left            =   240
            TabIndex        =   28
            Tag             =   "NoFileSysPage"
            Top             =   3480
            Width           =   3390
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Device Manager Page "
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
            Height          =   240
            Index           =   33
            Left            =   0
            TabIndex        =   27
            Tag             =   "NoDevMgrPage"
            Top             =   2730
            Width           =   3390
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Remove File menu from Explorer"
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
            Height          =   240
            Index           =   32
            Left            =   0
            TabIndex        =   26
            Tag             =   "NoFileMenu"
            Top             =   2490
            Width           =   3390
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Remove the Tildes in Short Filenames ""~"""
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
            Height          =   240
            Index           =   31
            Left            =   0
            TabIndex        =   25
            Tag             =   "NameNumericTail"
            Top             =   2250
            Width           =   3390
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Explorer's default context menu "
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
            Height          =   240
            Index           =   25
            Left            =   0
            TabIndex        =   24
            Tag             =   "NoViewContextMenu"
            Top             =   1995
            Width           =   3390
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Don't Save Settings at Exit "
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
            Height          =   240
            Index           =   24
            Left            =   0
            TabIndex        =   23
            Tag             =   "NoSaveSettings"
            Top             =   1755
            Width           =   3405
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Control Panel"
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
            Height          =   240
            Index           =   21
            Left            =   0
            TabIndex        =   22
            Tag             =   "NoControlPanel"
            Top             =   1035
            Width           =   3405
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Remove Username from the Start Menu"
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
            Height          =   240
            Index           =   22
            Left            =   0
            TabIndex        =   21
            Tag             =   "NoUserNameInStartMenu"
            Top             =   1275
            Width           =   3405
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Command Prompt and Batch Files"
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
            Height          =   240
            Index           =   23
            Left            =   0
            TabIndex        =   20
            Tag             =   "DisableCMD"
            Top             =   1515
            Width           =   3405
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable System Tray "
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
            Height          =   240
            Index           =   20
            Left            =   0
            TabIndex        =   19
            Tag             =   "NoTrayItemsDisplay"
            Top             =   795
            Width           =   3390
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Context Menus For the Taskbar"
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
            Height          =   240
            Index           =   19
            Left            =   0
            TabIndex        =   18
            Tag             =   "NoTrayContextMenu"
            Top             =   555
            Width           =   3390
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Network Neighborhood Icon"
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
            Height          =   240
            Index           =   18
            Left            =   0
            TabIndex        =   17
            Tag             =   "NoNetHood"
            Top             =   315
            Width           =   3390
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable the Shut Down Command"
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
            Height          =   240
            Index           =   17
            Left            =   0
            TabIndex        =   16
            Tag             =   "NoClose"
            Top             =   75
            Width           =   3390
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1290
         Left            =   240
         ScaleHeight     =   1290
         ScaleWidth      =   3690
         TabIndex        =   9
         Top             =   5160
         Width           =   3690
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable the Tools / Internet Options menu"
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
            Height          =   240
            Index           =   30
            Left            =   75
            TabIndex        =   14
            Tag             =   "NoBrowserOptions"
            Top             =   1035
            Width           =   3405
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable of selecting a download directory"
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
            Height          =   240
            Index           =   29
            Left            =   75
            TabIndex        =   13
            Tag             =   "NoBrowserOptions"
            Top             =   795
            Width           =   3555
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable the Tools / Internet Options menu"
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
            Height          =   240
            Index           =   28
            Left            =   75
            TabIndex        =   12
            Tag             =   "NoBrowserOptions"
            Top             =   555
            Width           =   3405
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable right-click context menu"
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
            Height          =   240
            Index           =   27
            Left            =   75
            TabIndex        =   11
            Tag             =   "NoBrowserContextMenu"
            Top             =   315
            Width           =   3405
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable option of closing Internet Explorer"
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
            Height          =   240
            Index           =   26
            Left            =   75
            TabIndex        =   10
            Tag             =   "NoBrowserClose"
            Top             =   75
            Width           =   3540
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   240
         ScaleHeight     =   1065
         ScaleWidth      =   3690
         TabIndex        =   4
         Top             =   4200
         Width           =   3690
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Display Settings Page "
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
            Height          =   240
            Index           =   11
            Left            =   75
            TabIndex        =   8
            Tag             =   "NoDispSettingsPage"
            Top             =   795
            Width           =   3405
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Screen Saver Settings Page "
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
            Height          =   240
            Index           =   10
            Left            =   75
            TabIndex        =   7
            Tag             =   "NoDispScrSavPage"
            Top             =   555
            Width           =   3405
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Display Background Page "
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
            Height          =   240
            Index           =   9
            Left            =   75
            TabIndex        =   6
            Tag             =   "NoDispBackgroundPage"
            Top             =   315
            Width           =   3405
         End
         Begin VB.CheckBox chkSystem 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hide the Display Appearance Page "
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
            Height          =   240
            Index           =   8
            Left            =   75
            TabIndex        =   5
            Tag             =   "NoDispAppearancePage"
            Top             =   75
            Width           =   3405
         End
      End
      Begin VB.CommandButton cmdTweak 
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   4080
         TabIndex        =   3
         Tag             =   "Apply tweak settings"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmdTweak 
         Caption         =   "Cek All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   4080
         TabIndex        =   2
         Tag             =   "Select all of tweak."
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmdTweak 
         Caption         =   "Clear All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   1
         Tag             =   "Clear all of tweak."
         Top             =   5520
         Width           =   1575
      End
      Begin Mizano.Abutton Abutton5 
         Height          =   375
         Left            =   5880
         TabIndex        =   44
         Top             =   4920
         Width           =   2295
         _ExtentX        =   4048
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
      Begin Mizano.Abutton Abutton3 
         Height          =   375
         Left            =   5880
         TabIndex        =   45
         Top             =   5400
         Width           =   2295
         _ExtentX        =   4048
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
      Begin Mizano.ProgressBar PBFix 
         Height          =   255
         Left            =   4080
         TabIndex        =   46
         Top             =   6120
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         Color           =   12937777
         Color2          =   12937777
      End
      Begin VB.Label LblFixed 
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
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   6960
         Width           =   5535
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   3240
         Picture         =   "frmRegistryFixer.frx":20D85
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Registry Fix and Setting"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3960
         TabIndex        =   48
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status  Registry Fixer  :"
         Height          =   255
         Left            =   4080
         TabIndex        =   47
         Top             =   5880
         Width           =   2415
      End
   End
   Begin VB.Image imageAbout 
      Height          =   735
      Left            =   240
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Image imageHome 
      Height          =   855
      Left            =   240
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Image imageMinimize 
      Height          =   855
      Left            =   17280
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imageMaximize 
      Height          =   855
      Left            =   18480
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imageClose 
      Height          =   855
      Left            =   19560
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmRegistryFixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imageAbout_Click()
 About.Show
End Sub

Private Sub imageClose_Click()
    Unload Me
End Sub

Private Sub imageHome_Click()
 Unload Me
    frmHome.Show
End Sub

Private Sub imageMaximize_Click()
    frmParent.WindowState = vbMaximized
End Sub

Private Sub imageMinimize_Click()
    frmParent.WindowState = vbMinimized
End Sub

