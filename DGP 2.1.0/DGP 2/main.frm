VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dynamic Graphics Plotter 1.0"
   ClientHeight    =   10935
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   15105
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   15105
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   5
      Left            =   9360
      TabIndex        =   37
      Top             =   6120
      Width           =   5655
      Begin VB.CommandButton cmddetailed 
         Caption         =   "detailed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   4080
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblhistory 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblhistory 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   44
         Top             =   720
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblhistory 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   43
         Top             =   1200
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblhistory 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   42
         Top             =   1680
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblhistory 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   41
         Top             =   2160
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblhistory 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   40
         Top             =   2640
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblhistory 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   39
         Top             =   3120
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblhistory 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   38
         Top             =   3600
         Visible         =   0   'False
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Index           =   4
      Left            =   9360
      TabIndex        =   26
      Top             =   2640
      Width           =   5655
      Begin VB.PictureBox graphcolor 
         AutoRedraw      =   -1  'True
         Height          =   375
         Left            =   3840
         ScaleHeight     =   315
         ScaleWidth      =   1635
         TabIndex        =   56
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton chktrace 
         Caption         =   "TRACE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdscale 
         Caption         =   "43"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   105
         Picture         =   "main.frx":0442
         TabIndex        =   36
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton cmdscale 
         Caption         =   "65"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   720
         Picture         =   "main.frx":0884
         TabIndex        =   35
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton cmdscale 
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1220
         Picture         =   "main.frx":0CC6
         TabIndex        =   34
         Top             =   2040
         Width           =   600
      End
      Begin VB.CommandButton cmdscale 
         Caption         =   "56"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   0
         Left            =   720
         Picture         =   "main.frx":1108
         TabIndex        =   33
         Top             =   1410
         Width           =   495
      End
      Begin VB.CommandButton cmdquality 
         Caption         =   "QUALITY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   32
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdcolor 
         Caption         =   "COLOR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1980
         TabIndex        =   31
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdcalc 
         Caption         =   "CALCULATOR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkscale 
         Height          =   495
         Left            =   720
         Picture         =   "main.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2040
         Width           =   495
      End
      Begin VB.CheckBox chkeraser 
         Caption         =   "ERASER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chklabel 
         Caption         =   "ADD LABEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.PictureBox screen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   120
      MouseIcon       =   "main.frx":169C
      ScaleHeight     =   8595
      ScaleWidth      =   9075
      TabIndex        =   4
      Top             =   2040
      Width           =   9135
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         ScaleHeight     =   435
         ScaleWidth      =   315
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   8640
         Top             =   7560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSScriptControlCtl.ScriptControl ScriptControl1 
         Left            =   8520
         Top             =   8040
         _ExtentX        =   1005
         _ExtentY        =   1005
         AllowUI         =   -1  'True
      End
      Begin VB.Shape Shape1 
         Height          =   135
         Left            =   0
         Shape           =   2  'Oval
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Index           =   3
      Left            =   9360
      TabIndex        =   3
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdrefresh 
         Caption         =   "REFESH"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         TabIndex        =   55
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtticks 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   54
         Text            =   "20"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtticks 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   52
         Text            =   "20"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtymax2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3720
         TabIndex        =   23
         Text            =   "10"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtymin2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Text            =   "-10"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtxmax2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3720
         TabIndex        =   20
         Text            =   "10"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtxmin2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   19
         Text            =   "-10"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtymax1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4800
         TabIndex        =   17
         Text            =   "10"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtymin1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3720
         TabIndex        =   16
         Text            =   "-10"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtxmax1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Text            =   "10"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtxmin1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Text            =   "-10"
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Optlim3 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Optlim2 
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   870
         Width           =   255
      End
      Begin VB.OptionButton Optlim1 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Number of ticks fo x axis                for y axis "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   53
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Label4 
         Caption         =   "Y from                          to"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "X from                          to"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "X from                   to                      Y from                  to"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Index           =   2
      Left            =   6840
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdclrall 
         Caption         =   "Clear all"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdpaper 
         Caption         =   "Graph paper"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblycoord 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblxcoord 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1140
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdprint 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdplot 
         Caption         =   "PLOT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtinput 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   840
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "main.frx":17EE
         Left            =   120
         List            =   "main.frx":17F8
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label lbltype 
         Caption         =   "Y(x)="
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ListBox List1 
         Height          =   1425
         ItemData        =   "main.frx":1808
         Left            =   120
         List            =   "main.frx":182A
         TabIndex        =   47
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnusaveas 
         Caption         =   "&Save as..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
         Shortcut        =   +{F12}
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "&Options"
      Begin VB.Menu mnugrid 
         Caption         =   "Grid"
         Begin VB.Menu mnugridon 
            Caption         =   "On"
            Shortcut        =   ^O
         End
         Begin VB.Menu mnugridoff 
            Caption         =   "Off"
            Checked         =   -1  'True
            Shortcut        =   ^F
         End
      End
      Begin VB.Menu mnuaddlabel 
         Caption         =   "A&dd label"
      End
      Begin VB.Menu mnucalculator 
         Caption         =   "&Calculator..."
      End
      Begin VB.Menu mnucolor 
         Caption         =   "C&olor..."
      End
      Begin VB.Menu mnuerase 
         Caption         =   "&Eraser"
      End
      Begin VB.Menu mnugraph 
         Caption         =   "G&raph paper..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnurefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnutrace 
         Caption         =   "&Trace..."
      End
      Begin VB.Menu mnuquality 
         Caption         =   "&Quality..."
      End
      Begin VB.Menu mnudetailed 
         Caption         =   "Detailed history..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnusetting 
      Caption         =   "&Settings"
      Begin VB.Menu mnuautocorrection 
         Caption         =   "&Autocorrection"
         Begin VB.Menu mnuautoon 
            Caption         =   "On"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuautooff 
            Caption         =   "Off"
         End
      End
   End
   Begin VB.Menu mnuaction 
      Caption         =   "Actio&n"
      Begin VB.Menu mnuplot 
         Caption         =   "&Plot"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuclear 
         Caption         =   "Clear all"
      End
   End
   Begin VB.Menu mnuscales 
      Caption         =   "&Scales"
      Begin VB.Menu mnuextend 
         Caption         =   "Extend"
         Begin VB.Menu mnuey 
            Caption         =   "in Y axes"
         End
         Begin VB.Menu mnuex 
            Caption         =   "in X axes"
         End
      End
      Begin VB.Menu mnusqueez 
         Caption         =   "Squeeze"
         Begin VB.Menu mnusy 
            Caption         =   "in Y axes"
         End
         Begin VB.Menu mnusx 
            Caption         =   "in X axes"
         End
      End
      Begin VB.Menu mnuzoombox 
         Caption         =   "&Zoom box"
      End
   End
   Begin VB.Menu mnulang 
      Caption         =   "&Language"
      Begin VB.Menu mnuenglish 
         Caption         =   "English"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnufrench 
         Caption         =   "Français"
      End
      Begin VB.Menu mnudutch 
         Caption         =   "Nederlands"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelp1 
         Caption         =   "Contents..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnucontact 
         Caption         =   "contact..."
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "about..."
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rcol, gcol, bcol As Long         'RGB values for color of graph
Dim autocorrect As Boolean           'flag for checking scale limits
Dim xstart, ystart As Single         'first point for plotting square on main window
Dim xprevious As Single              'second point for plotting...
Dim yprevious As Single              'square on main window
Dim historycount, tracecount As Byte 'counters for and arrays for history and trace
Dim xstep, ystep As Single           'steps for ticks on axis
Dim xsize, ysize As Single           'total scale sizex for x and y
Dim xaxisval, yaxisval As Single     'limits for plotting axis
Dim ystep1, xstep1 As Single         'number of ticks on axis y and x
Dim ytick, xtick As Integer          'used to avoid collision
Dim linestepx, linestepy As Single   'width of ticks
Dim linecount As Single              'counter for plotting ticks
Dim xminchk, xmaxchk As Single       'used to cleare screen for next plotted graph
Dim yminchk, ymaxchk As Single       'if the scales are different
Dim gridmode As Boolean              'used to varify whether to plot grid or remove it


Private Sub chktrace_Click()
For allcount = 0 To 7
    If tracehist(allcount) <> "" Then
    trace.List1.AddItem tracehist(allcount)
    End If
Next
tracer = True
chkeraser = 0
chklabel = 0
chktrace = 1
trace.Show vbModal 'keep form on top
End Sub
Private Sub cmdclrall_Click()
clrnum = 0
screen.Cls
For allcount = 0 To 7
tracehist(allcount) = ""
lblhistory(allcount).Caption = ""
lblhistory(allcount).Visible = False
historycount = 0
Next
trace.List1.Clear
End Sub
Rem this changes color using
Rem a common dialog

Private Sub cmdcolor_Click()
    Dim Cl As OLE_COLOR
    On Error Resume Next
    CD1.Flags = 0 ' set flags to 0
    CD1.DialogTitle = "Choose Color for your next graph" 'caption
    CD1.ShowColor ' show th window
    graphcolor.BackColor = CD1.color 'set color
End Sub
Private Sub cmdpaper_Click()
gpaper1.Show vbModal
End Sub
Private Sub cmdplot_Click()
screen.Refresh
If txtinput <> "" Then ' if textbox contains equation then...
On Error GoTo errors
addtrace ' add to the list of equations to trace
calclimits ' calculte min and max limits
screen.DrawWidth = 1 ' width of the graph
ScriptControl1.language = "VBScript" ' set script language to VBScript
screen.ForeColor = graphcolor.BackColor
cmddetailed.Visible = True ' make detail option enabled
mnudetailed.Enabled = True ' on the first run
estring = LCase(txtinput.Text) ' assign the contents of textbox to string estring
SqueezeSpaces estring ' call function SqueezeSpaces to delete any space between constants and variables
txtinput.Text = estring ' assign squeezed text back to equation text box
nleftbrackets = NumOccStr(estring, "(") 'get the number of leftbrackets
nrightbrackets = NumOccStr(estring, ")") 'get the number of right brackets
If nleftbrackets <> nrightbrackets Then '   if number of left brackets is not
    If nleftbrackets > nrightbrackets Then 'equal to number of right brackets
        Select Case langval
        Case 1
        MsgBox ("Not enough of ')'")           ' show error and exit sub
        Case 2
        MsgBox ("Niet voldoende ')'")
        Case 3
        MsgBox ("Pas assez ')' ")
        End Select
    Else
        Select Case langval
        Case 1
        MsgBox ("Not enough of '('")
        Case 2
        MsgBox ("Niet voldoende '('")
        Case 3
        MsgBox ("Pas assez '(' ")
        End Select
    End If
Exit Sub
End If
If lbltype.Caption = "Y(x)=" Then ' get the type of a function
plot txtinput.Text, True ' pass an equation to the plot function, with function type
Else: plot txtinput.Text, False
End If
history 'manage history
plotaxis 'plot ticks and axis
End If
nextcol ' make next color randomly choosen
Exit Sub
errors: ' this will handle error by displaying what the error is
res& = MsgBox("Error #" & CStr(Err.Number) & " " & Err.Description + vbCrLf + cstring$, vbOKOnly)
On Error GoTo 0
Exit Sub
End Sub

Sub nextcol()
rcol = (colorsrnd(clrnum) And 255)
gcol = (colorsrnd(clrnum) And 65280) / 256
bcol = (colorsrnd(clrnum) And 16711680) / 65536
graphcolor.BackColor = RGB(rcol, gcol, bcol)
clrnum = clrnum + 1
If clrnum = 8 Then clrnum = 0
End Sub
Private Sub cmdPrint_Click()
Dim sWide As Single, sTall As Single
Dim rv As Long
' I have printed accidenetely many times, so decided to add this msgbox
print1& = MsgBox("Are you sure you want to print your main window?", vbYesNo)
If print1& = vbYes Then
screen.SetFocus ' Set focus on the main PictureBox ie. where the plotting and labels are
      Picture2.AutoRedraw = True
      rv = SendMessage(screen.hwnd, WM_PAINT, Picture2.hDC, 0) '  take a shot of the screen
      rv = SendMessage(screen.hwnd, WM_PRINT, Picture2.hDC, _
      PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
      Picture2.AutoRedraw = True
      Picture2.DrawWidth = 5
       For allcount = 0 To 7 ' print functions if present
        If lblhistory(allcount).Caption <> "" Then
        Picture2.ForeColor = lblhistory(allcount).BackColor
        Picture2.Print lblhistory(allcount).Caption
        End If
       Next
      Picture2.Picture = Picture2.Image 'put the picture into image, which will be printed
      Printer.PrintQuality = 300 ' set quality,300 is supported by all printers
      Printer.Print ""
      ' use the printer Paintpicture
      Printer.PaintPicture Picture2.Picture, 1100, 0
      Printer.EndDoc 'end printing
End If
End Sub

Private Sub cmdquality_Click()
frmOptions.Show vbModal
End Sub



Private Sub cmdrefresh_Click()
calclimits 'calc limits
screen.Cls
plotall
End Sub
Rem change avaliable scaling options depending on the type of equations
Private Sub Combo1_click()
lbltype.Caption = Combo1.List(Combo1.ListIndex) & "="
If lbltype.Caption = "Y(x)=" Then
Optlim3.Enabled = False
Optlim2.Enabled = True
Else
Optlim3.Enabled = True
Optlim2.Enabled = False
End If
txtinput.Text = ""
If Optlim2.Enabled = False And Optlim2 = True Then Optlim3 = True
If Optlim3.Enabled = False And Optlim3 = True Then Optlim2 = True
End Sub

Private Sub cmdcalc_Click()
calculator.Show vbModal 'show calculator
End Sub
Rem load default settings
Private Sub Form_Load()
'resol 'enable if cannot have resolution higher than 800x600
clr
nextcol 'show random color
autocorrect = True ' set autocorrect option to true
Optlim1 = True
Combo1.Text = "Y(x)"
drawline = False
update_settings ' settings for the quality
Picture2.Width = screen.Width 'set picture to print equal to the size of actual main picture screen
Picture2.Height = screen.Height
main.Caption = "Dynamic Graphics Plotter 1.0"
langval = 1
End Sub
Rem this is calculating  x or y minimum and maximum
Sub calclimits()
ScriptControl1.language = "VBScript"
If Optlim1 = True Then ' if all the limits specifyed
    'check whether lower limits are lower than upper limits
    If Val(txtxmin1.Text) > Val(txtxmax1.Text) Or Val(txtymin1.Text) > Val(txtymax1.Text) Then
        If autocorrect = False Then ' if autocorrect is off then
MsgBox ("The value of upper limits should be bigger than the value of lower limits") 'display error
        Else ' if autocorrect is on then correct error
            If Val(txtxmin1.Text) > Val(txtxmax1.Text) Then ' correct error for x limits
            tempval = Val(txtxmin1.Text)  ' assign value of min limit to temperary variable
            txtxmin1.Text = txtxmax1.Text ' assign max value to min value
            txtxmax1.Text = tempval       ' assign min value to max from temp variable
            End If
            If Val(txtymin1.Text) > Val(txtymax1.Text) Then ' correct error for y limits
            tempval = Val(txtymin1.Text)  ' assign value of min limit to temperary variable
            txtymin1.Text = txtymax1.Text ' assign max value to min value
            txtymax1.Text = tempval       ' assign min value to max from temp variable
            End If
        End If
    End If
xmin = Val(txtxmin1.Text): xmax = Val(txtxmax1.Text) ' set values of xmin,xmax,ymin and ymax
ymin = Val(txtymin1.Text): ymax = Val(txtymax1.Text) 'be equal to values of text boxes
ElseIf Optlim2 = True Then 'if only limits for x axis are specifyed
If Val(txtxmin2.Text) > Val(txtxmax2.Text) Then ' check if max value > then min value
    If autocorrect = False Then ' if autoreect is off then display error
    MsgBox "The value of upper limits should be bigger than the value of lower limits", vbOKOnly
    Else ' if autocorrect is on then swap min value with max value
    tempval = Val(txtxmin2.Text)
    txtxmin2.Text = txtxmax2.Text
    txtxmax2.Text = tempval
    End If
End If
xmin = Val(txtxmin2.Text): xmax = Val(txtxmax2.Text) 'assign values of textboxes to variables
'calculate values of ymin and ymax
ymin = 999: ymax = -999 'set values to extreme,which is impossible
ScriptControl1.ExecuteStatement ("Dim x") 'assign x variable to script control
ScriptControl1.ExecuteStatement ("x = " & xmin) 'let x to be min value
For x = xmin To xmax Step 0.1 ' going from min value to max value with step 0.1(accurate enough)
Y = ScriptControl1.Eval(txtinput.Text) ' evaluate function
On Error Resume Next ' error hardly possible because was already checked
If ymin > Y Then ymin = Y ' if Y >,< than ymax,ymin assign value of Y to ymax,ymin
If ymax < Y Then ymax = Y ' this will defenetly happen on first run
ScriptControl1.ExecuteStatement ("x = x + 0.1") ' increase value of x in script control by step
Next x 'loop
ElseIf Optlim3 = True Then 'if only limits for y axis are specifyed
'check whether lower limits are lower than upper limits
If Val(txtymin2.Text) > Val(txtymax2.Text) Then
    If autocorrect = False Then ' if autocorrect is off then display error
    MsgBox "The value of upper limits should be bigger than the value of lower limits", vbOKOnly
    Else ' if autocorrect is on then correct error
    tempval = Val(txtymin2.Text)
    txtymin2.Text = txtymax2.Text
    txtymax2.Text = tempval
    End If
End If
ymin = Val(txtymin2.Text): ymax = Val(txtymax2.Text)
xmin = 999: xmax = -999 'set values to extreme
ScriptControl1.ExecuteStatement ("Dim y")
ScriptControl1.ExecuteStatement ("y = " & ymin)
For Y = ymin To ymax Step 0.1 ' calculate values for xmin and xmax
x = ScriptControl1.Eval(txtinput.Text)
If xmin > x Then xmin = x
If xmax < x Then xmax = x
ScriptControl1.ExecuteStatement ("y = y + 0.1")
Next Y
End If
checklimits
End Sub
Rem the following function plots axis and ticks
Sub plotaxis()
If txtticks(0).Text < 0 Or txtticks(0).Text > 50 Then _
txtticks(0).Text = 20 ' do not allow more than 50 and less than 0 ticks
If txtticks(1).Text < 0 Or txtticks(1).Text > 50 Then _
txtticks(1).Text = 20
ytick = 0: xtick = 0
ysize = Abs(screen.ScaleHeight) ' get the full size
xsize = Abs(screen.ScaleWidth)
If ymin < 0 And ymax >= 0 Then 'ie standard (dont need to move y axis ticks)
yaxisval = 0
ytick = 1
ElseIf ymin >= 0 Then ' need to move y axis ticks to the left
yaxisval = ymin
ytick = 2
ElseIf ymax < 0 Then 'need to move y axis ticks to the right
yaxisval = ymax
ytick = 3
End If
If xmin <= 0 And xmax > 0 Then ' same as with y axis but for x axis
xaxisval = 0: xtick = 1
ElseIf xmin > 0 Then
xaxisval = xmin: xtick = 2
ElseIf xmax <= 0 Then
xaxisval = xmax: xtick = 3
End If
screen.DrawWidth = 2 ' make axis wider than the plot
screen.ForeColor = vbBlack ' color always black
screen.Line (xmin, yaxisval)-(xmax, yaxisval) 'plot the x axis
screen.Line (xaxisval, ymin)-(xaxisval, ymax) 'plot the y axis
linestepx = xaxisval + ((xmax - xmin) / 400) 'width of ticks on x axis
linestepy = yaxisval + ((ymax - ymin) / 300) '                  y axis
ystep1 = ysize / Val(txtticks(1).Text) ' number of ticks
For linecount = ymin To ymax Step ystep1
Rem the following code is plotting ticks,
Rem it plotes differently for different situstions
Rem so that user will always see the ticks and the axis
Rem depending on scale it changes coordinates at
Rem which to plot ticks and print values
Select Case xtick
Case 1
screen.Line (-linestepx, linecount)-(linestepx, linecount)
Case 2
screen.Line (xmin, linecount)-(linestepx, linecount)
screen.CurrentX = xmin + xsize / 100
Case 3
screen.Line (xmax, linecount)-(xmax - xsize / 400, linecount)
screen.CurrentX = xmax - xsize / 25
End Select
screen.Print Roundnum(linecount, 2) ' print raunded value
Next                                ' so that the values do not mess up
xstep1 = xsize / Val(txtticks(0))
For linecount = xmin To xmax Step xstep1
Select Case ytick
Case 1
screen.Line (linecount, -linestepy)-(linecount, linestepy)
screen.CurrentY = -0.05
Case 2
screen.Line (linecount, ymin)-(linecount, linestepy)
screen.CurrentY = ymin + ysize / 35
Case 3
screen.Line (linecount, ymax)-(linecount, ymax - ysize / 400)
screen.CurrentY = xmax - xsize / 100
End Select
If linecount <> 0 Then screen.Print Roundnum(linecount, 2)
Next
End Sub
Sub checklimits()
'the purpose of this is to clear the screen if scales change
If Optlim1 = True Then ' if all scales specifyed then check for changes
    If lblhistory(0).Caption <> "" Then 'if anything changed then clear screen
        If xminchk <> Val(txtxmin1.Text) Or xmaxchk <> Val(txtxmax1.Text) Or _
        yminchk <> Val(txtymin1.Text) Or ymaxchk <> Val(txtymax1.Text) Then
        screen.Cls
        Call cmdclrall_Click
        cmddetailed.Visible = False 'disable detail option
        mnudetailed.Enabled = False
        End If
    End If
xminchk = Val(txtxmin1.Text): xmaxchk = Val(txtxmax1.Text)
yminchk = Val(txtymin1.Text): ymaxchk = Val(txtymax1.Text)
ElseIf Optlim2 = True Then
        screen.Cls
        Call cmdclrall_Click
xminchk = Val(txtxmin2.Text): xmaxchk = Val(txtxmax2.Text)
Else
        screen.Cls
        Call cmdclrall_Click
yminchk = Val(txtymin2.Text): ymaxchk = Val(txtymax2.Text)
End If

End Sub
Private Sub graphcolor_Click()
nextcol ' choose another color on click
End Sub

Private Sub lblhistory_Click(Index As Integer)
histstack = Index
If eqhistorytype(Index) = True Then
corrtypefunc = True
Else: corrtypefunc = False
End If
corrfunc = equationhistory(Index)
details.txtcorrinput.Text = equationhistory(Index)
details.Show vbModal
End Sub

Private Sub List1_Click()
txtinput.Text = txtinput.Text & List1.List(List1.ListIndex) ' add function to listbox
End Sub

Private Sub mnusaveas_Click()
CD1.CancelError = True
On Error Resume Next
CD1.FileName = "graph" ' file name to save
CD1.DialogTitle = "Save To Bitmap" ' caption
CD1.Filter = "BitMap File|*.bmp" ' type of format
CD1.DefaultExt = ".bmp"
CD1.ShowSave 'show save window
SavePicture screen.Image, CD1.FileName ' actual saving
 Exit Sub
End Sub

Private Sub screen_MouseDown(Button As Integer, Shift _
As Integer, x As Single, Y As Single)
If chkscale = 1 Then ' if zoom box choosen
screen.DrawWidth = 1 ' in case if drawwidth bigger
    xmin = x ' get xmin coordinate
    ymin = Y ' get xmax coordinate
    If Button = 1 Then ' if button on the mouse is clicked
        xstart = x  'First Point
        ystart = Y
        xprevious = xstart 'Second Point
        yprevious = ystart
        screen.AutoRedraw = False
    End If
End If
screen.CurrentX = x ' change current coordinates
screen.CurrentY = Y
If chklabel = 1 Then 'if add label is chosen
screen.ForeColor = vbBlack
Select Case langval
Case 1
screen.Print (InputBox("please enter the text to input")) ' invitation to input label
Case 2
screen.Print (InputBox("Gelieve de invoertekst in te geven"))
Case 3
screen.Print (InputBox("Veillez donner"))
End Select
End If
End Sub

Private Sub screen_mousemove(Button As Integer, _
Shift As Integer, x As Single, Y As Single)
 If lblxcoord.BackColor = &HC0FFFF Then ' depending on the label color
    lblxcoord = "x=" & Round(Str(x), 2) ' display rounded coordinates for x
 Else: lblxcoord = "x=" & Str(x) 'display coordinates for x
 End If
 If lblycoord.BackColor = &HC0FFFF Then ' same as for x but for y
    lblycoord = "y=" & Round(Str(Y), 2)
  Else: lblycoord = "y=" & Str(x)
 End If
If chkscale = 1 Then
If Button <> 1 Then Exit Sub
    screen.Refresh
    screen.Line (xstart, ystart)-(x, Y), , B 'Draws Square
End If
If chkeraser = 1 Then 'if eraser is chosen
If Button = 1 Then
screen.Line (screen.CurrentX, screen.CurrentY)-(x, Y) ' drawing line with foreclor=backcolor
End If                                                ' to turn everything what goes over into backcolor
End If
End Sub
Private Sub screen_MouseUp(Button As Integer, _
Shift As Integer, x As Single, Y As Single)
screen.AutoRedraw = True
Dim autoset As Boolean ' declare autoset, to automatically replace max with min
If chkscale = 1 Then   ' points if autocorrection is off
screen.Cls
chkscale = 0
xmax = x ' get x max
ymax = Y ' get y max
If autocorrect = False Then ' replacement if autocorrection is off
autoset = True
autocorrect = True
End If
Optlim1 = True
Call plotall
If autoset = True Then autocorrect = False
End If
End Sub
Rem managing history of plotted graphs
Sub history()
equationhistory(historycount) = txtinput.Text ' assign text
If lbltype.Caption = "Y(x)=" Then
eqhistorytype(historycount) = True ' history type needed for detailed equation plotting
Else: eqhistorytype(historycount) = False
End If
'historydet.hisequation(historycount).Caption = lbltype.Caption & txtinput.Text
lblhistory(historycount).Caption = lbltype.Caption & txtinput.Text
lblhistory(historycount).BackColor = screen.ForeColor
lblhistory(historycount).Visible = True ' make label visible if has equation in it
lblhistory(historycount).FontBold = True ' set font to bold
lblhistory(historycount).FontSize = 10
If historycount < 7 Then ' raise historycount for next equation
historycount = historycount + 1
Else: historycount = 0 ' 0 historycount if all 8 equations are filled up
End If
End Sub
Rem this procedure is adding plotted equations to a trace window
Sub addtrace()
If lbltype.Caption = "Y(x)=" Then
tracex(tracecount) = True ' tracex dealing with type of equation
Else: tracex(tracecount) = False
End If
traceequation(tracecount) = txtinput.Text ' add equation
If tracecount < 7 Then
tracecount = tracecount + 1
Else: tracecount = 0 'can hold no more than 8 equations
End If
tracehist(tracecount) = txtinput.Text
End Sub
'Read from .ini file, original code by Richard Hayden ( PlotX) , you can find plotX at psc
' Special thanx for introducing me to MS script control, and code for .ini Richard
Rem this procedure is updating quality setings, reading it from .ini file
Sub update_settings()
Rem depending on the number, setting accuracy
    If GetIniInfo(App.Path & "\dgp.ini", "GRAPH DRAWING", "Accuracy", "3") = "1" Then
        accuracy = 0.1
    ElseIf GetIniInfo(App.Path & "\dgp.ini", "GRAPH DRAWING", "Accuracy", "3") = "2" Then
        accuracy = 0.01
    ElseIf GetIniInfo(App.Path & "\dgp.ini", "GRAPH DRAWING", "Accuracy", "3") = "3" Then
        accuracy = 0.001
    ElseIf GetIniInfo(App.Path & "\dgp.ini", "GRAPH DRAWING", "Accuracy", "3") = "4" Then
        accuracy = 0.0005
    Else
        accuracy = 0.001 ' in case if someone changed contents of the file, to keep program working
    End If
    Exit Sub
End Sub
Rem eraser function which used to turn anything what mouse goes over into background color
Private Sub chkeraser_Click()
chklabel = 0 ' canceling checked buttons
chktrace = 0
If chkeraser = 1 Then
screen.MousePointer = 99 ' changing mouse pointer
screen.ForeColor = screen.BackColor
screen.DrawWidth = 18
Else
screen.MousePointer = 0 ' when canceled
screen.ForeColor = vbBlack
screen.DrawWidth = 1
End If
End Sub

Private Sub chklabel_Click()
chkeraser = 0 'canceling other buttons
chktrace = 0
End Sub
Rem This function is responsible for plotting graphics in detailed window
Rem it plots the graphics which are stored in history
Private Sub cmddetailed_Click()
For allcount = 0 To 7
historydet.hisequation(allcount).Caption = _
lblhistory(allcount).Caption
historydet.screenhis(allcount).Scale (-10, 10)-(10, -10) ' setting scales
historydet.screenhis(allcount).Print allcount + 1 ' printing number according to chkbuttons on detailed window
historydet.screenhis(allcount).Line (-10, 0)-(10, 0) ' drawing lines
historydet.screenhis(allcount).Line (0, -10)-(0, 10)
Next
For allcount = 0 To 7
If historydet.hisequation(allcount) <> "" Then
    ScriptControl1.Reset
    ScriptControl1.language = "VBScript"
    If eqhistorytype(allcount) = True Then ' if graph type Y(x) then declare x, do evaluations
        ScriptControl1.ExecuteStatement ("Dim x")
        ScriptControl1.ExecuteStatement ("x = " & -10)
        historydet.screenhis(allcount).CurrentX = -10
        historydet.screenhis(allcount).CurrentY = ScriptControl1.Eval(equationhistory(allcount))
        For x = xmin To xmax Step 0.01 ' step gives good quality and satisfying speed
        Y = ScriptControl1.Eval(equationhistory(allcount))
        On Error Resume Next
        If Y >= ymin And Y <= ymax Then
        historydet.screenhis(allcount).Line (historydet.screenhis(allcount).CurrentX, _
        historydet.screenhis(allcount).CurrentY)-(x, Y) ' plot all the lines in detailed window
        historydet.screenhis(allcount).CurrentX = x     ' using lines
        historydet.screenhis(allcount).CurrentY = Y
        End If
        ScriptControl1.ExecuteStatement ("x = x + " & 0.01)
        Next
    Else
    ScriptControl1.ExecuteStatement ("Dim y") ' if type is X(y) then do following
    ScriptControl1.ExecuteStatement ("y = " & -10)
    historydet.screenhis(allcount).CurrentY = -10
    historydet.screenhis(allcount).CurrentX = ScriptControl1.Eval(equationhistory(allcount))
    On Error Resume Next
    For Y = ymin To ymax Step 0.01
        x = ScriptControl1.Eval(equationhistory(allcount))
        'If X >= xmin And X <= xmax Then
        historydet.screenhis(allcount).Line (historydet.screenhis(allcount).CurrentX, _
        historydet.screenhis(allcount).CurrentY)-(x, Y) ' same thing as with Y(x)
        historydet.screenhis(allcount).CurrentX = x
        historydet.screenhis(allcount).CurrentY = Y
        'End If
        ScriptControl1.ExecuteStatement ("y = y + " & 0.01)
    Next
    End If
End If
Next
historydet.Show vbModal ' show the actual window
End Sub
Rem this functio extends or squezes graphic in either y or x axis
Private Sub cmdscale_Click(Index As Integer)
Select Case Index
Case 0 ' extend in y axis
If ymin < 0 Then
ymin = ymin * 2
Else: ymin = ymin - ymin
End If
If ymax > 0 Then
ymax = ymax * 2
Else: ymax = ymax - ymax
End If
Case 1 ' extend in x axis
If xmin < 0 Then
xmin = xmin * 2
Else: xmin = xmin - xmin
End If
If xmax > 0 Then
xmax = xmax * 2
Else: xmax = xmax - xmax
End If
Case 2 'squeeze in y axis
If ymin < 0 Then
ymin = ymin / 2
Else: ymin = ymin + ymin
End If
If ymax > 0 Then
ymax = ymax / 2
Else: ymax = ymax + ymax
End If
Case 3 'squeeze in x axis
If xmin < 0 Then
xmin = xmin / 2
Else: xmin = xmin + xmin
End If
If xmax > 0 Then
xmax = xmax / 2
Else: xmax = xmax + xmax
End If
End Select
screen.Cls
Call plotall ' replots the graph
End Sub
Rem this function reploting the graph, used to change number of ticks
Rem and for squeezing, extending axis
Sub plotall()
If xmin > xmax Then tempval = xmin: xmin = xmax: xmax = tempval ' if max value<min value
If ymin > ymax Then tempval = ymin: ymin = ymax: ymax = tempval ' then it reverses it
screen.DrawWidth = 1
For allcount = 0 To 7
If equationhistory(allcount) = "" Then
Else ' if there is equation it replots it
    screen.ForeColor = lblhistory(allcount).BackColor
    If eqhistorytype(allcount) = True Then
    plot equationhistory(allcount), True
    Else: plot equationhistory(allcount), False
    End If
End If
Next
screen.ForeColor = vbBlack
plotaxis ' plotting axis
End Sub
Rem below the functions which perfomed by menu buttons
Rem most of them are colling allready made procedures
Private Sub mnuautooff_Click()
autocorrect = False
mnuautoon.Checked = False
mnuautooff.Checked = True
End Sub

Private Sub mnuautoon_Click()
autocorrect = True
mnuautooff.Checked = False
mnuautoon.Checked = True
End Sub

Private Sub mnuExit_Click()
End
End Sub
Rem this function plotting grid to make it easier
Rem orientating on graph
Private Sub mnugridon_Click()
If gridmode = False Then ' if grid on is pressed it sets color to (130,130,70)
mnugridon.Checked = True ' deal with checking buttons
mnugridoff.Checked = False
screen.ForeColor = RGB(130, 130, 70)
Else: screen.ForeColor = screen.BackColor ' if grid off is pressed it sets color equal to background
End If                                    ' will draw the same lines but vertually erase existent
screen.DrawWidth = 1
screen.DrawStyle = 2
For allcount = ymin To ymax Step ystep1
screen.Line (xmin, allcount)-(xmax, allcount) ' draws lines from xmin to xmax
Next                                          ' with distance already specifyed
For allcount = xmin To xmax Step xstep1       ' draws lines from ymin to ymax
screen.Line (allcount, ymin)-(allcount, ymax)
Next
gridmode = False
screen.ForeColor = vbBlack
If mnugridoff.Checked = True Then plotaxis
End Sub
Private Sub mnugridoff_Click() ' call mnugrid_on with gridmode equal=true
mnugridon.Checked = False      ' so that means that color of lines will be
mnugridoff.Checked = True      ' equal to backgroung color
gridmode = True
Call mnugridon_Click
End Sub
Private Sub lblycoord_Click()
If lblycoord.BackColor = &HC0FFFF Then ' changes color and so changes modes between
lblycoord.BackColor = &HC0FFC0         ' rounded value and not rounded
Else: lblycoord.BackColor = &HC0FFFF
End If
End Sub
Private Sub lblxcoord_Click()
If lblxcoord.BackColor = &HC0FFFF Then
lblxcoord.BackColor = &HC0FFC0
Else: lblxcoord.BackColor = &HC0FFFF
End If
End Sub
Sub form_unload(cancel As Integer)
End ' close all windows if cross button pressed
End Sub
' menu actions
Private Sub mnuaddlabel_Click()
chklabel = 1
End Sub
Private Sub mnucalculator_Click()
calculator.Show
End Sub
Private Sub mnucolor_Click()
Call cmdcolor_Click
End Sub
Private Sub mnuerase_Click()
chkeraser = 1
End Sub
Private Sub mnugraph_Click()
Gpaper.Show
End Sub
Private Sub mnutrace_Click()
trace.Show
End Sub
Private Sub mnuquality_Click()
frmOptions.Show
End Sub
Private Sub mnudetailed_Click()
Call cmddetailed_Click
End Sub
Private Sub mnuplot_Click()
Call cmdplot_Click
End Sub
Private Sub mnuPrint_Click()
Call cmdPrint_Click
End Sub
Private Sub mnuclear_Click()
Call cmdclrall_Click
End Sub
Private Sub mnuzoombox_Click()
chkscale = 1
End Sub
Private Sub mnuex_Click()
Call cmdscale_Click(1)
End Sub

Private Sub mnuey_Click()
Call cmdscale_Click(0)
End Sub
Private Sub mnusy_Click()
Call cmdscale_Click(2)
End Sub
Private Sub mnusx_Click()
Call cmdscale_Click(3)
End Sub
Private Sub mnurefresh_Click()
Call cmdrefresh_Click
End Sub
Private Sub mnuabout_Click()
frmAbout.Show vbModal
End Sub
Private Sub mnucontact_Click()
' simple command to open default browser with my adress and topic "DGP"
Shell "Start.exe " & "mailto:cyberjet@cyberarmy.com?Subject=DGP", 0
On Error Resume Next
End Sub
Private Sub mnuhelp1_Click()
CD1.HelpFile = "DGP1.hlp" ' use common dialog to open help file
CD1.HelpCommand = cdlHelpContents
CD1.ShowHelp
End Sub
Private Sub mnuenglish_Click()
langval = 1
language (langval)
End Sub
Private Sub mnudutch_Click()
langval = 2
language (langval)
End Sub
Private Sub mnufrench_Click()
langval = 3
language (langval)
End Sub
Sub resol()
screen.Height = 6135
screen.Width = 6855
List1.Width = 1095
Frame1(0).Width = 1335
Frame1(1).Left = 1560
Frame1(1).Width = 2895
Combo1.Width = 2655
txtinput.Width = 2055
cmdplot.Width = 1215
cmdPrint.Width = 1215
cmdPrint.Left = 1500
Frame1(4).Left = 7080
Frame1(4).Width = 4455
Frame1(2).Left = 4550
Frame1(3).Left = 7100
chktrace.Top = 1200
chktrace.Left = 1980
cmdquality.Top = 1680
cmdquality.Left = 1980
graphcolor.Top = 2160
graphcolor.Left = 1980
Label2.Caption = "X from           to            Y from          to"
txtxmin1.Left = 1200
txtxmin1.Width = 375
txtxmax1.Left = 1920
txtxmax1.Width = 375
txtymin1.Left = 2880
txtymin1.Width = 375
txtymax1.Left = 3480
txtymax1.Width = 375
Label3.Left = 600
Label4.Left = 600
txtxmin2.Left = 1320
txtxmax2.Left = 2640
txtymin2.Left = 1320
txtymax2.Left = 2640
Label1.Caption = "Number of ticks X                   Y"
txtticks(0).Left = 1800
txtticks(1).Left = 2760
Label2.Width = 3375
Label3.Width = 2895
Label4.Width = 3375
Label1.Width = 3375
Frame1(3).Width = 4455
cmdrefresh.Left = 3480
cmdrefresh.Height = 855
Frame1(5).Left = 7080
cmddetailed.Top = 1680
lblhistory(3).Top = 4080
End Sub
