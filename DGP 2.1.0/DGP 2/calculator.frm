VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form calculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   4785
   ClientLeft      =   9360
   ClientTop       =   5850
   ClientWidth     =   5100
   Icon            =   "calculator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5100
   Begin VB.TextBox txtxval 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4320
      TabIndex        =   33
      Text            =   "1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.OptionButton optdeg 
      Caption         =   "Degrees"
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   1200
      Width           =   975
   End
   Begin VB.OptionButton optrad 
      Caption         =   "Radians"
      Height          =   195
      Left            =   240
      TabIndex        =   31
      Top             =   840
      Width           =   975
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   2280
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.CommandButton cmdeval 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Evaluate"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdplace 
      Caption         =   "Place"
      Height          =   255
      Left            =   3960
      TabIndex        =   26
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtanswer 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   2640
      TabIndex        =   13
      Top             =   2040
      Width           =   2175
      Begin VB.CommandButton cmdrgtbracket 
         Caption         =   ")"
         Height          =   495
         Left            =   1560
         TabIndex        =   29
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdlftbracket 
         Caption         =   "("
         Height          =   495
         Left            =   840
         TabIndex        =   28
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdsqr 
         Caption         =   "Sqr"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmddevide 
         Caption         =   "/"
         Height          =   495
         Left            =   1560
         TabIndex        =   22
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmdmmultiply 
         Caption         =   "*"
         Height          =   495
         Left            =   840
         TabIndex        =   21
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmdminus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmdplus 
         Caption         =   "+"
         Height          =   495
         Left            =   1560
         TabIndex        =   19
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdpower2 
         Caption         =   "X^2"
         Height          =   495
         Left            =   840
         TabIndex        =   18
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdpower 
         Caption         =   "^"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdtan 
         Caption         =   "Tan"
         Height          =   495
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdcos 
         Caption         =   "Cos"
         Height          =   495
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdsin 
         Caption         =   "Sin"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
      Begin VB.CommandButton cmdabs 
         Caption         =   "abs"
         Height          =   495
         Left            =   840
         TabIndex        =   30
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmddot 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   12
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmddigit 
         Caption         =   "9"
         Height          =   495
         Index           =   9
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmddigit 
         Caption         =   "8"
         Height          =   495
         Index           =   8
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmddigit 
         Caption         =   "7"
         Height          =   495
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmddigit 
         Caption         =   "6"
         Height          =   495
         Index           =   6
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmddigit 
         Caption         =   "5"
         Height          =   495
         Index           =   5
         Left            =   840
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmddigit 
         Caption         =   "4"
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmddigit 
         Caption         =   "3"
         Height          =   495
         Index           =   3
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmddigit 
         Caption         =   "2"
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   4
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmddigit 
         Caption         =   "1"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmddigit 
         Caption         =   "0"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   495
      End
   End
   Begin VB.TextBox txtdisplay 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "X="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   34
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim answer As Double
Rem the Sub's below are used to add assigned values
Private Sub cmdabs_Click()
txtdisplay.Text = txtdisplay.Text & "abs("
End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdcos_Click()
txtdisplay.Text = txtdisplay.Text & "Cos("
End Sub

Private Sub cmddevide_Click()
txtdisplay.Text = txtdisplay.Text & "/"
End Sub

Private Sub cmddigit_Click(Index As Integer)
txtdisplay.Text = txtdisplay.Text + cmddigit(Index).Caption
End Sub

Private Sub cmddot_Click()
txtdisplay.Text = txtdisplay.Text & "."
End Sub

Private Sub cmdlftbracket_Click()
txtdisplay.Text = txtdisplay.Text & "("
End Sub

Private Sub cmdminus_Click()
txtdisplay.Text = txtdisplay.Text & "-"
End Sub

Private Sub cmdmmultiply_Click()
txtdisplay.Text = txtdisplay.Text & "*"
End Sub
Private Sub cmdplace_Click()
main.txtinput.Text = main.txtinput.Text & txtanswer.Text
End Sub
Private Sub cmdplus_Click()
txtdisplay.Text = txtdisplay.Text & "+"
End Sub

Private Sub cmdpower_Click()
txtdisplay.Text = txtdisplay.Text & "^"
End Sub

Private Sub cmdpower2_Click()
txtdisplay.Text = txtdisplay.Text & "^2"
End Sub

Private Sub cmdrgtbracket_Click()
txtdisplay.Text = txtdisplay.Text & ")"
End Sub

Private Sub cmdsign_Click()
If Val(txtdisplay.Text) < 0 Then
txtdisplay.Text = Abs(Val(txtdisplay.Text))
Else: txtdisplay.Text = "-" & txtdisplay.Text
End If
End Sub

Private Sub cmdsin_Click()
txtdisplay.Text = txtdisplay.Text & "Sin("
End Sub

Private Sub cmdsqr_Click()
txtdisplay.Text = txtdisplay.Text & "Sqr("
End Sub

Private Sub cmdtan_Click()
txtdisplay.Text = txtdisplay.Text & "Tan("
End Sub
' evaluating contents of textbox
Private Sub Cmdeval_Click()
pival = "Const pi=3.14159265358979" 'adding value of pi
ScriptControl1.AddCode (pival)
ScriptControl1.ExecuteStatement ("Dim x")
ScriptControl1.ExecuteStatement ("x=" & txtxval.Text)
If optrad = True Then ' answer in radians
answer = ScriptControl1.Eval(txtdisplay.Text)
Else ' convert answer into degrees
answer = (ScriptControl1.Eval(txtdisplay.Text)) * 180 / 3.1415927
End If
txtanswer.Text = answer
txtdisplay.Text = ""
End Sub

Private Sub Form_Load()
optrad = True ' default mode is radians
End Sub
