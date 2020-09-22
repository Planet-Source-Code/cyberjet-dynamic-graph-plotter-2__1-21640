VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form trace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tracer"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "trace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.FlatScrollBar tracescrl 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   2
      Arrows          =   65536
      Orientation     =   8323073
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   2
      Top             =   1590
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose equation to trace"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label lbltracey 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lbltracex 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "trace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim equation As String
Dim tracevalx, tracevaly As Double
Private Sub cmdok_Click()
Unload Me
tracer = False
main.chktrace = False
main.Shape1.Visible = False
End Sub


Private Sub form_load()
Me.Left = main.Frame1(4).Left
End Sub

Private Sub List1_Click()
equation = traceequation(List1.ListIndex)
tracescrl.Min = xmin * 10
tracescrl.Max = xmax * 10
main.Shape1.Visible = True
End Sub
Rem with the change of scrool bar calculate the points
Rem and move shape on calculated coordinates
Private Sub tracescrl_Change()
main.ScriptControl1.Reset
main.ScriptControl1.language = "VBScript"
If tracex(List1.ListIndex) = True Then
main.ScriptControl1.ExecuteStatement ("Dim x")
main.ScriptControl1.ExecuteStatement ("x=" & tracescrl.Value / 10)
X = tracescrl.Value / 10
Y = main.ScriptControl1.Eval(equation)
Else
main.ScriptControl1.ExecuteStatement ("Dim y")
main.ScriptControl1.ExecuteStatement ("y=" & tracescrl.Value / 10)
Y = tracescrl.Value / 10
X = main.ScriptControl1.Eval(equation)
End If
main.Shape1.Left = X - (main.Shape1.Width / 2)
lbltracex.Caption = X
main.Shape1.Top = Y + (main.Shape1.Height / 2)
lbltracey.Caption = Y
End Sub
Sub form_unload(cancel As Integer)
'if cross is pressed make sure chktrace is unchecked
main.chktrace = False
End Sub
