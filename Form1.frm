VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Basic Super Object Property Setting Utility Thingy - Because Programming is meant to be easy"
   ClientHeight    =   6750
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame13 
      Caption         =   "Objects found in the form"
      Height          =   4695
      Left            =   120
      TabIndex        =   78
      Top             =   0
      Width           =   4575
      Begin VB.ListBox List1 
         Height          =   4335
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   79
         Top             =   240
         Width           =   4335
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   4800
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   12632256
      TabCaption(0)   =   "Properties"
      TabPicture(0)   =   "Form1.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkTrimAdd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSave"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Key Verification"
      TabPicture(1)   =   "Form1.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdGenKeyV"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkElse(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkElse(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Text Verification"
      TabPicture(2)   =   "Form1.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command4"
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(2)=   "Command3"
      Tab(2).Control(3)=   "Frame9"
      Tab(2).Control(4)=   "Frame8"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Focus"
      TabPicture(3)   =   "Form1.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdGetIndex"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame11"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdClearFocus"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "List2"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdMoveDown"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdMoveUp"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cmdRemove"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "cmdAddtoList"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label3"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "About"
      TabPicture(4)   =   "Form1.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Timer1"
      Tab(4).Control(1)=   "cmdPlayControl"
      Tab(4).Control(2)=   "Picture1"
      Tab(4).Control(3)=   "Label5"
      Tab(4).Control(4)=   "Label4"
      Tab(4).ControlCount=   5
      Begin VB.CommandButton Command4 
         Caption         =   "Generate ""Select on GotFocus"" Events"
         Height          =   375
         Left            =   -74760
         TabIndex        =   87
         Top             =   4440
         Width           =   4815
      End
      Begin VB.CommandButton cmdGetIndex 
         Caption         =   ">> by Index"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   77
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   120
         TabIndex        =   72
         Top             =   2520
         Width           =   5175
         Begin VB.OptionButton optTxtText 
            Caption         =   "Null String = """""
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.TextBox txtTextText 
            Height          =   285
            Left            =   3480
            TabIndex        =   74
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkTxtText 
            Caption         =   "Text.Text"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   0
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.OptionButton optTxtText 
            Caption         =   "Expression"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   76
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame12 
         Height          =   615
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Width           =   5175
         Begin VB.TextBox txtTextLocked 
            Height          =   285
            Left            =   3480
            TabIndex        =   71
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkTxtLocked 
            Caption         =   "Text.Locked"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton optTxtLocked 
            Caption         =   "Expression"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   69
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optTxtLocked 
            Caption         =   "FALSE"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optTxtLocked 
            Caption         =   "TRUE"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   67
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -70320
         Top             =   5160
      End
      Begin VB.CommandButton cmdPlayControl 
         Caption         =   "Pause"
         Height          =   375
         Left            =   -72840
         TabIndex        =   61
         Top             =   5040
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   3735
         Left            =   -74880
         ScaleHeight     =   245
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   341
         TabIndex        =   60
         Top             =   600
         Width           =   5175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Generate Code"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73560
         TabIndex        =   59
         Top             =   5280
         Width           =   2535
      End
      Begin VB.Frame Frame11 
         Caption         =   "Options"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   55
         Top             =   3720
         Width           =   5175
         Begin VB.CheckBox Check1 
            Caption         =   "On GotFocus select all text"
            Height          =   255
            Left            =   360
            TabIndex        =   58
            Top             =   960
            Width           =   3135
         End
         Begin VB.OptionButton optTextFocus 
            Caption         =   "Move on Cursor > Text Length"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   57
            Top             =   480
            Width           =   2775
         End
         Begin VB.OptionButton optTextFocus 
            Caption         =   "Move on Enter"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   56
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.CommandButton cmdClearFocus 
         Caption         =   "Clear"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   54
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Frame Frame10 
         Caption         =   "Options"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   49
         Top             =   2400
         Width           =   2535
         Begin VB.CheckBox Check2 
            Caption         =   "Set focus if..."
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   1200
            Width           =   2295
         End
         Begin VB.OptionButton optVerify 
            Caption         =   "Verify all with OR"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   53
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Trim Text"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton optVerify 
            Caption         =   "Verify all with AND"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton optVerify 
            Caption         =   "Verify each seperately"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.ListBox List2 
         Height          =   2790
         Left            =   -73440
         TabIndex        =   47
         Top             =   720
         Width           =   3615
      End
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "Move Down"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   46
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "Move Up"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   45
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<<"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   44
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddtoList 
         Caption         =   ">>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   43
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox chkElse 
         Caption         =   "Allow everything except"
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   42
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CheckBox chkElse 
         Caption         =   "Allow only..."
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   41
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Generate "
         Height          =   375
         Left            =   -72240
         TabIndex        =   38
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Frame Frame9 
         Caption         =   "Test Property With"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   33
         Top             =   1200
         Width           =   5175
         Begin VB.OptionButton optTestWith 
            Caption         =   "Value"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTestWith 
            Caption         =   "Variable"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtVerVal 
            Height          =   285
            Left            =   1200
            TabIndex        =   35
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtVarVal 
            Height          =   285
            Left            =   1200
            TabIndex        =   34
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Test Property If"
         Height          =   615
         Left            =   -74880
         TabIndex        =   30
         Top             =   480
         Width           =   5175
         Begin VB.OptionButton optEq 
            Caption         =   "Equal to"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Tag             =   "="
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optEq 
            Caption         =   "Not Equal to"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   31
            Tag             =   "<>"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdGenKeyV 
         Caption         =   "Generate Key Verification"
         Height          =   375
         Left            =   -74760
         TabIndex        =   17
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Frame Frame7 
         Caption         =   "Don't Allow"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   24
         Top             =   2640
         Width           =   5055
         Begin VB.CheckBox chkDelete 
            Caption         =   "Delete"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   93
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtOtherKeys 
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   91
            Top             =   1320
            Width           =   4575
         End
         Begin VB.CheckBox chkOthers 
            Caption         =   "Others"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   89
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox chkEnter 
            Caption         =   "Enter"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   40
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox chkBksp 
            Caption         =   "Backspace"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   29
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox chkEsc 
            Caption         =   "Escape"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   28
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkNum 
            Caption         =   "Numbers"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkLcase 
            Caption         =   "Lowercase"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkUcase 
            Caption         =   "Uppercase"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Allow"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   18
         Top             =   600
         Width           =   5055
         Begin VB.CheckBox chkDelete 
            Caption         =   "Delete"
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   92
            Top             =   720
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtOtherKeys 
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   90
            Top             =   1320
            Width           =   4575
         End
         Begin VB.CheckBox chkOthers 
            Caption         =   "Others"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   88
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox chkEnter 
            Caption         =   "Enter"
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   39
            Top             =   960
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkBksp 
            Caption         =   "Backspace"
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   23
            Top             =   480
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkEsc 
            Caption         =   "Escape"
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkNum 
            Caption         =   "Numbers"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkLcase 
            Caption         =   "Lowercase"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox chkUcase 
            Caption         =   "Uppercase"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Generate Properties"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get Text Values"
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Set Button.Enabled to"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   5175
         Begin VB.OptionButton optBtnEnabled 
            Caption         =   "TRUE"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optBtnEnabled 
            Caption         =   "FALSE"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optBtnEnabled 
            Caption         =   "Expression"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtBtnvar 
            Height          =   285
            Left            =   3480
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Set ComboBox Property"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   5175
         Begin VB.OptionButton optCbo 
            Caption         =   ".Text = """""
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optCbo 
            Caption         =   ".ListIndex = -1"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   5175
         Begin VB.TextBox txtTextEnabled 
            Height          =   285
            Left            =   3480
            TabIndex        =   64
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkTxtEnabled 
            Caption         =   "Text.Enabled"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton optTxtEnabled 
            Caption         =   "FALSE"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optTxtEnabled 
            Caption         =   "TRUE"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optTxtEnabled 
            Caption         =   "Expression"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   65
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkTrimAdd 
         Caption         =   "Trim All Text"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Use the Trim() function on all text properties"
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "EnderSoft ©2003  http://www.geocities.com/saintender.geo"
         Height          =   255
         Left            =   -74760
         TabIndex        =   63
         Top             =   4680
         Width           =   4935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Another useless programming utility brought to you by"
         Height          =   255
         Left            =   -74880
         TabIndex        =   62
         Top             =   4440
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "SetFoucus() order"
         Height          =   375
         Left            =   -73440
         TabIndex        =   48
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   4575
      Begin VB.OptionButton optSelectType 
         Caption         =   "Inverse"
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   85
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optSelectType 
         Caption         =   "None"
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   84
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optSelectType 
         Caption         =   "All"
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   83
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton optSelectType 
         Caption         =   "Combo Boxes"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   82
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optSelectType 
         Caption         =   "Text Boxes"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   81
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optSelectType 
         Caption         =   "Command Buttons"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   80
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOPen 
         Caption         =   "&Open Form..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020                            ' (DWORD) dest = source

Private Type VBOBJ
    CName As String
    CType As String
    CIndex As Long
End Type

Dim FocusList(200) As VBOBJ
Dim MyObj(200) As VBOBJ
Dim showabout As Boolean
Dim py As Long
Dim abtmsg(11) As String

Private Sub chkElse_Click(Index As Integer)
    chkUcase(1 - Index).Enabled = 1 - chkElse(Index).Value
    chkLcase(1 - Index).Enabled = 1 - chkElse(Index).Value
    chkNum(1 - Index).Enabled = 1 - chkElse(Index).Value
    chkDelete(1 - Index).Enabled = 1 - chkElse(Index).Value
    chkEsc(1 - Index).Enabled = 1 - chkElse(Index).Value
    chkBksp(1 - Index).Enabled = 1 - chkElse(Index).Value
    chkEnter(1 - Index).Enabled = 1 - chkElse(Index).Value
End Sub

Private Sub cmdAddtoList_Click()
    lastcount = List2.ListCount
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            FocusList(i + lastcount) = MyObj(i)
            List2.AddItem MyObj(i).CName & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "")
        End If
    Next
End Sub

Private Sub cmdGenKeyV_Click()
Dim i As Integer
Dim indexed As String
Dim CaseAllow As String
Dim CaseNAllow As String
Dim temp As String

    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            Select Case MyObj(i).CType
            Case "TextBox"
                CaseAllow = ""
                CaseNAllow = ""

                indexed = ""

                If MyObj(i).CIndex >= 0 Then indexed = "Index as Integer,"

                temp = temp & "Private Sub " & Trim(MyObj(i).CName) & "_KeyPress(" & indexed & "KeyAscii as Integer)" & vbCrLf
                temp = temp & vbTab & "Select Case KeyAscii" & vbCrLf

                If chkUcase(0).Value = 1 Then
                    CaseAllow = " Asc(" & Chr(34) & "A" & Chr(34) & ") to Asc(" & Chr(34) & "Z" & Chr(34) & ")"
                End If

                If chkLcase(0).Value = 1 Then
                    If CaseAllow <> "" Then CaseAllow = CaseAllow & ","
                    CaseAllow = CaseAllow & " Asc(" & Chr(34) & "a" & Chr(34) & ") to Asc(" & Chr(34) & "z" & Chr(34) & ")"
                End If

                If chkNum(0).Value = 1 Then
                    If CaseAllow <> "" Then CaseAllow = CaseAllow & ","
                    CaseAllow = CaseAllow & " Asc(" & Chr(34) & "0" & Chr(34) & ") to Asc(" & Chr(34) & "9" & Chr(34) & ")"
                End If

                If chkEsc(0).Value = 1 Then
                    If CaseAllow <> "" Then CaseAllow = CaseAllow & ","
                    CaseAllow = CaseAllow & " vbKeyEscape"
                End If

                If chkBksp(0).Value = 1 Then
                    If CaseAllow <> "" Then CaseAllow = CaseAllow & ","
                    CaseAllow = CaseAllow & " vbKeyBack"
                End If

                If chkBksp(0).Value = 1 Then
                    If CaseAllow <> "" Then CaseAllow = CaseAllow & ","
                    CaseAllow = CaseAllow & " vbKeyDelete"
                End If

                If chkEnter(0).Value = 1 Then
                    If CaseAllow <> "" Then CaseAllow = CaseAllow & ","
                    CaseAllow = CaseAllow & " vbKeyReturn"
                End If

                If chkOthers(0).Value = 1 Then
                    For j = 1 To Len(txtOtherKeys(0).text)
                        If CaseAllow <> "" Then CaseAllow = CaseAllow & ","
                        CaseAllow = CaseAllow & " Asc(" & Chr(34) & Mid(txtOtherKeys(0).text, j, 1) & Chr(34) & ")"
                    Next
                End If

                If chkElse(1).Value = 0 Then
                    temp = temp & vbTab & "Case " & CaseAllow & vbCrLf
                    temp = temp & vbTab & vbTab & "'Allow These Keys" & vbCrLf
                End If

                If chkElse(0).Value = 1 Then
                    temp = temp & vbTab & "Case Else" & vbCrLf
                    temp = temp & vbTab & vbTab & "' Do Not Allow All Other Keys" & vbCrLf
                    temp = temp & vbTab & vbTab & "KeyAscii = 0" & vbCrLf
                Else

                    If chkUcase(1).Value = 1 Then
                        CaseNAllow = CaseNAllow & " Asc(" & Chr(34) & "A" & Chr(34) & ") to Asc(" & Chr(34) & "Z" & Chr(34) & ")"
                    End If

                    If chkLcase(1).Value = 1 Then
                        If CaseNAllow <> "" Then CaseNAllow = CaseNAllow & ","
                        CaseNAllow = CaseNAllow & " Asc(" & Chr(34) & "a" & Chr(34) & ") to Asc(" & Chr(34) & "z" & Chr(34) & ")"
                    End If

                    If chkNum(1).Value = 1 Then
                        If CaseNAllow <> "" Then CaseNAllow = CaseNAllow & ","
                        CaseNAllow = CaseNAllow & " Asc(" & Chr(34) & "0" & Chr(34) & ") to Asc(" & Chr(34) & "9" & Chr(34) & ")"
                    End If

                    If chkEsc(1).Value = 1 Then
                        If CaseNAllow <> "" Then CaseNAllow = CaseNAllow & ","
                        CaseNAllow = CaseNAllow & " vbKeyEscape"
                    End If

                    If chkBksp(1).Value = 1 Then
                        If CaseNAllow <> "" Then CaseNAllow = CaseNAllow & ","
                        CaseNAllow = CaseNAllow & " vbKeyBack"
                    End If

                    If chkEnter(1).Value = 1 Then
                        If CaseNAllow <> "" Then CaseNAllow = CaseNAllow & ","
                        CaseNAllow = CaseNAllow & " vbKeyReturn"
                    End If

                    If chkOthers(1).Value = 1 Then
                        For j = 1 To Len(txtOtherKeys(1).text)
                            If CaseNAllow <> "" Then CaseNAllow = CaseNAllow & ","
                            CaseNAllow = CaseNAllow & " Asc(" & Chr(34) & Mid(txtOtherKeys(1).text, j, 1) & Chr(34) & ")"
                        Next
                    End If

                    temp = temp & vbTab & "Case " & CaseNAllow & vbCrLf
                    temp = temp & vbTab & vbTab & "' Do Not Allow These Keys" & vbCrLf
                    temp = temp & vbTab & vbTab & "KeyAscii = 0" & vbCrLf

                End If

                If chkElse(1).Value = 1 Then
                    temp = temp & vbTab & "Case Else" & vbCrLf
                    temp = temp & vbTab & vbTab & "' Allow All Other Keys" & vbCrLf
                End If

                temp = temp & vbTab & "End Select" & vbCrLf
                temp = temp & "End Sub" & vbCrLf & vbCrLf & vbCrLf
            End Select
        End If
    Next

    CopyCode temp

End Sub

Private Sub cmdInvert_Click()

End Sub

Private Sub cmdMoveDown_Click()
    m = List2.ListIndex
    q = List2.List(m)
    If m < 0 Or m = List2.ListCount - 1 Then Exit Sub
    List2.RemoveItem (m)
    List2.AddItem q, m + 1
    List2.Selected(m + 1) = True


End Sub

Private Sub cmdMoveUp_Click()
    m = List2.ListIndex
    q = List2.List(m)
    If m <= 0 Then Exit Sub
    List2.RemoveItem (m)
    List2.AddItem q, m - 1
    List2.Selected(m - 1) = True

End Sub

Private Sub cmdPlayControl_Click()
    If cmdPlayControl.Caption = "Pause" Then
        cmdPlayControl.Caption = "Play"
    ElseIf cmdPlayControl.Caption = "Play" Then
        cmdPlayControl.Caption = "Pause"
    End If
End Sub

Private Sub cmdRemove_Click()
    m = List2.ListIndex
    If m < 0 Then Exit Sub
    List2.RemoveItem (m)
    For i = m To List2.ListCount - 2
        FocusList(i) = FocusList(i + 1)
    Next
End Sub

Private Sub cmdSave_Click()
Dim i As Integer
    outtext = ""
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            Select Case MyObj(i).CType
            Case "TextBox"
                If chkTxtText.Value = 1 Then AddText IIf(chkTrimAdd.Value = 1, "Trim(", "") & Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Text" & IIf(chkTrimAdd.Value = 1, ")", "") & " = " & Chr$(34) & Chr$(34)

                If chkTxtLocked.Value = 1 Then
                    If optTxtLocked(0).Value = True Then AddText Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Locked = FALSE"
                    If optTxtLocked(1).Value = True Then AddText Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Locked = TRUE"
                End If

                If chkTxtEnabled.Value = 1 Then
                    If optTxtEnabled(0).Value = True Then AddText Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Enabled = FALSE"
                    If optTxtEnabled(1).Value = True Then AddText Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Enabled = TRUE"
                    If optTxtEnabled(2).Value = True Then AddText Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Enabled = " & txtTextEnabled.text
                End If

            Case "ComboBox"
                If optCbo(0).Value = True Then AddText IIf(chkTrimAdd.Value = 1, "Trim(", "") & Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Text" & IIf(chkTrimAdd.Value = 1, ")", "") & " = " & Chr$(34) & Chr$(34)
                If optCbo(1).Value = True Then AddText Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".ListIndex = -1"

            Case "CommandButton"
                If optBtnEnabled(0).Value = True Then AddText Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Enabled = FALSE"
                If optBtnEnabled(1).Value = True Then AddText Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Enabled = TRUE"
                If optBtnEnabled(2).Value = True Then AddText Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Enabled = " & Trim(txtBtnvar.text)
            End Select
        End If
    Next
    Close 1

    CopyCode outtext
End Sub





Private Sub Command1_Click()
Dim i As Integer
    outtext = ""
    
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            Select Case MyObj(i).CType
            Case "TextBox"
                If chkTxtText.Value = 1 Then AddText " = " & IIf(chkTrimAdd.Value = 1, "Trim(", "") & Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Text" & IIf(chkTrimAdd.Value = 1, ")", "")

                '     If chkTxtLocked.Value = 1 Then
                '      If optTxtLocked(0).Value = True Then Print #1, Trim(myobj(i).cname) & ".Locked = FALSE"
                '      If optTxtLocked(1).Value = True Then Print #1, Trim(myobj(i).cname) & ".Locked = TRUE"
                '     End If

                '     If chkTxtEnabled.Value = 1 Then
                '      If optTxtEnabled(0).Value = True Then Print #1, Trim(myobj(i).cname) & ".Enabled = FALSE"
                '      If optTxtEnabled(1).Value = True Then Print #1, Trim(myobj(i).cname) & ".Enabled = TRUE"
                '     End If

            Case "ComboBox"
                AddText " = " & Trim(MyObj(i).CName) & IIf(MyObj(i).CIndex > -1, "(" & MyObj(i).CIndex & ")", "") & ".Text"

                '    Case "CommandButton"
                '     If optBtnEnabled(0).Value = True Then Print #1, Trim(myobj(i).cname) & ".Enabled = FALSE"
                '     If optBtnEnabled(1).Value = True Then Print #1, Trim(myobj(i).cname) & ".Enabled = TRUE"
                '     If optBtnEnabled(2).Value = True Then Print #1, Trim(myobj(i).cname) & ".Enabled = " & Trim(txtBtnvar.Text)
            End Select

        End If
    Next
    Close 1

    CopyCode outtext
End Sub

Private Sub Command3_Click()
Dim temp As String
Dim i As Integer

    optr = IIf(optEq(0).Value, " = ", IIf(optEq(1).Value, " <> ", ""))
    comp = IIf(optTestWith(0).Value, Chr(34) & txtVerVal.text & Chr(34), IIf(optTestWith(1).Value, txtVarVal.text, ""))
    oper = IIf(optVerify(0).Value, " And ", IIf(optVerify(1).Value, " Or ", ""))

    temp = "If "
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            Select Case MyObj(i).CType
            Case "TextBox", "ComboBox"
                If optVerify(2) Then
                    temp = temp & IIf(Check3.Value = 1, "Trim(", "") & MyObj(i).CName & ".Text " & IIf(Check3.Value = 1, ")", "") & optr & comp & " Then " & vbCrLf & IIf(Check2.Value = 1, vbTab & MyObj(i).CName & ".SetFocus", "") & vbCrLf & "End if" & vbCrLf & vbCrLf & "If "
                Else
                    temp = temp & IIf(Check3.Value = 1, "Trim(", "") & MyObj(i).CName & ".Text " & IIf(Check3.Value = 1, ")", "") & optr & comp & oper
                End If
            End Select
        End If
    Next
    temp = temp & " Then " & vbCrLf & "End if"

    CopyCode temp
End Sub

Private Sub Command4_Click()
Dim temp As String
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            Select Case MyObj(i).CType
            Case "TextBox", "ComboBox"
                temp = temp & "Public Sub " & MyObj(i).CName & "_GotFocus()" & vbCrLf
                temp = temp & vbTab & MyObj(i).CName & ".SelStart =0" & vbCrLf
                temp = temp & vbTab & MyObj(i).CName & ".SelLength = Len(" & MyObj(i).CName & ".Text)" & vbCrLf
                temp = temp & "End Sub" & vbCrLf & vbCrLf
            End Select
        End If
    Next
    CopyCode temp
End Sub

Public Sub CopyCode(code As String)
    frmOut.txtOut.text = code
    Clipboard.SetText code
    frmOut.Show vbModal
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim j As Integer

    py = Picture1.ScaleHeight
    abtmsg(0) = "©2003 - EnderSoft"
    abtmsg(1) = "-----------------"
    abtmsg(2) = ""
    abtmsg(3) = "Visual Basic Super Object Property Setting Utility Thingy "
    abtmsg(4) = "Because Programming is meant to be easy"
    abtmsg(5) = ""
    abtmsg(6) = "Greetz:"
    abtmsg(7) = " irc.lfx.org/#romhack"
    abtmsg(8) = "     Chojin, ProtoCat"
    abtmsg(9) = ""
    abtmsg(10) = " Galaxynet/#axn"
    abtmsg(11) = "      star|et,CMSA,Mizerab|e,fujitaka,Laydee,Kougaiji,teina,menova"
End Sub

Private Sub ReadFile(filename As String)
    List2.Clear
    List1.Clear
    filename = Replace(filename, "\\", "\")
    Open filename For Input As 1
    While Not EOF(1)
        Line Input #1, temp

        i = InStr(temp, "VB.TextBox")
        mtype = "TextBox"

        If i = 0 Then
            i = InStr(temp, "VB.ComboBox")
            mtype = "ComboBox"
        End If

        If i = 0 Then
            i = InStr(temp, "VB.CommandButton")
            mtype = "CommandButton"
        End If

        MyObj(j).CType = mtype
        If i > 0 Then
            ControlName = Mid(temp, i + Len("VB." & mtype) + 1, 1 + Len(temp) - i + Len("VB." & mtype))
            mindex = ""

            While InStr(temp, "End") = 0
                Line Input #1, temp
                i = InStr(temp, " Index")
                If i > 0 Then
                    mindex = Trim(Mid(temp, InStr(temp, "=") + 1, 7))
                End If
            Wend
            MyObj(j).CName = Trim(ControlName)
            MyObj(j).CIndex = IIf(mindex = "", -1, CLng(Val(mindex)))
            MyObj(j).CType = mtype

            List1.AddItem mtype & " - " & Trim(ControlName) & IIf(mindex = "", "", "(" & mindex & ")")
            j = j + 1

        End If
    Wend
    Close 1
End Sub

Private Sub Label1_Click()

End Sub

Private Sub mnuFileOPen_Click()
    On Error GoTo ERRCANCEL
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Form Files|*.frm|Text Files|*.txt|All Files|*.*"
    CommonDialog1.ShowOpen
    ReadFile CommonDialog1.filename
    Exit Sub
ERRCANCEL:
    If Err.Number <> cdlCancel Then
        MsgBox Err.Number & ": " & Err.Description, vbOKOnly, App.Title
    End If
End Sub

Private Sub optSelectType_Click(Index As Integer)
    For i = 0 To List1.ListCount - 1
        Select Case optSelectType(Index).Caption
        Case "Command Buttons"
            If MyObj(i).CType = "CommandButton" Then List1.Selected(i) = True
        Case "Text Boxes"
            If MyObj(i).CType = "TextBox" Then List1.Selected(i) = True
        Case "Combo Boxes"
            If MyObj(i).CType = "ComboBox" Then List1.Selected(i) = True
        Case "All"
            List1.Selected(i) = True
        Case "None"
            List1.Selected(i) = False
        Case "Inverse"
            List1.Selected(i) = Not List1.Selected(i)
        End Select
    Next
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.TabCaption(SSTab1.Tab) = "About" Then
        Timer1.Enabled = True
        py = Picture1.ScaleHeight
        Picture1.Cls
    Else
        Timer1.Enabled = False
    End If


End Sub

Private Sub Timer1_Timer()
    If cmdPlayControl.Caption = "Pause" Then py = py - 1
    If py < 0 - ((UBound(abtmsg) + 10) * 16) Then py = Picture1.ScaleHeight
    Picture1.Cls
    For i = 0 To UBound(abtmsg)
        If py + (i * 16) < -16 Then i = i + 1
        If py + (i * 16) > Picture1.ScaleHeight Then Exit For
        TextOut Picture1.hdc, 0, py + (i * 16), abtmsg(i), Len(abtmsg(i))
    Next

    Picture1.Refresh
    DoEvents
End Sub
