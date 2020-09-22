VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detail level"
   ClientHeight    =   4050
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   4650
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdapply 
      Caption         =   "APPLY"
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
      Left            =   2160
      TabIndex        =   17
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "CANCEL"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lines/Dots"
      Height          =   1335
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmddots 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdlines 
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Use Dots"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Use Lines"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Accuracy"
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton Option1 
         Caption         =   "Excellent quality but very slow"
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Good quality"
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Average quality,fast"
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Poor quality only suitalbe if using lines"
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   1695
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdapply_Click()
Dim strTemp As String
Dim intTemp As Integer
    Do While intTemp < 4        ' looking for value which is set to true
        intTemp = intTemp + 1   ' increasing by 1 up to 4
        If Option1(intTemp).Value = True Then ' rewrites the value into .ini file
            WriteIniInfo "GRAPH DRAWING", "Accuracy", CStr(intTemp), App.Path & "\dgp.ini"
            Exit Do
        End If
    Loop
    
    Frame1.Enabled = True
    Call main.update_settings ' call updating of settings
    Me.Hide
    Exit Sub
End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub



Private Sub cmdlines_Click()
cmdlines.BackColor = &HFF00& 'changing color
cmddots.BackColor = &HC0C0C0
drawline = True
End Sub

Private Sub Cmddots_Click()
cmddots.BackColor = &HFF00& 'changing color
cmdlines.BackColor = &HC0C0C0
drawline = False
End Sub

Private Sub Form_Load()
Dim strTemp As String
Dim intTemp As Integer
If drawline = True Then ' changes color depending on the enabled mode
cmddots.BackColor = &HC0C0C0
cmdlines.BackColor = &HFF00&
Else
cmdlines.BackColor = &HC0C0C0
cmddots.BackColor = &HFF00&
End If
strTemp = GetIniInfo(App.Path & "\dgp.ini", "GRAPH DRAWING", "Accuracy", "3") ' reading from ini file
intTemp = CInt(strTemp)
Option1(intTemp).Value = True ' changes enabled option depending on the value
'center the form
Me.Move (screen.Width - Me.Width) / 2, (screen.Height - Me.Height) / 2

End Sub





