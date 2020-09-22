VERSION 5.00
Begin VB.Form historydet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detailed history"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   Icon            =   "historydet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   13590
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000C0&
      Height          =   2295
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   7080
      Width           =   3255
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   27
      Top             =   8880
      Width           =   3255
   End
   Begin VB.CommandButton cmdcomb 
      Caption         =   "COMBINE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   26
      Top             =   7080
      Width           =   3255
   End
   Begin VB.CheckBox chkcomb 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6600
      Width           =   735
   End
   Begin VB.CheckBox chkcomb 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6600
      Width           =   735
   End
   Begin VB.CheckBox chkcomb 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6600
      Width           =   735
   End
   Begin VB.CheckBox chkcomb 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6600
      Width           =   735
   End
   Begin VB.CheckBox chkcomb 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6600
      Width           =   735
   End
   Begin VB.CheckBox chkcomb 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6600
      Width           =   735
   End
   Begin VB.CheckBox chkcomb 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6600
      Width           =   735
   End
   Begin VB.CheckBox chkcomb 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
      Width           =   735
   End
   Begin VB.PictureBox hiscomb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      Height          =   2775
      Left            =   10200
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   8
      Top             =   6600
      Width           =   3255
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "combined"
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox screenhis 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      Height          =   2775
      Index           =   7
      Left            =   10200
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   7
      Top             =   3360
      Width           =   3255
   End
   Begin VB.PictureBox screenhis 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      Height          =   2775
      Index           =   6
      Left            =   6840
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   6
      Top             =   3360
      Width           =   3255
   End
   Begin VB.PictureBox screenhis 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      Height          =   2775
      Index           =   5
      Left            =   3480
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   3360
      Width           =   3255
   End
   Begin VB.PictureBox screenhis 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      Height          =   2775
      Index           =   4
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   4
      Top             =   3360
      Width           =   3255
   End
   Begin VB.PictureBox screenhis 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      Height          =   2775
      Index           =   3
      Left            =   10200
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox screenhis 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      Height          =   2775
      Index           =   2
      Left            =   6840
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox screenhis 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      Height          =   2775
      Index           =   1
      Left            =   3480
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.PictureBox screenhis 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      Height          =   2775
      Index           =   0
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblcomb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10200
      TabIndex        =   17
      Top             =   9480
      Width           =   3255
   End
   Begin VB.Label hisequation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No graph here"
      Height          =   255
      Index           =   7
      Left            =   10200
      TabIndex        =   16
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label hisequation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No graph here"
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   15
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label hisequation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No graph here"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   14
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label hisequation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No graph here"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label hisequation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No graph here"
      Height          =   255
      Index           =   3
      Left            =   10200
      TabIndex        =   12
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label hisequation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No graph here"
      Height          =   255
      Index           =   2
      Left            =   6840
      TabIndex        =   11
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label hisequation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No graph here"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   10
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label hisequation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No graph here"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   3255
   End
End
Attribute VB_Name = "historydet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcomb_Click()
lblcomb.Caption = ""
hiscomb.Cls
hiscomb.Scale (-10, 10)-(10, -10) 'specify scale and draw lines
hiscomb.Line (-10, 0)-(10, 0)
hiscomb.Line (0, -10)-(0, 10)
For allcount = 0 To 7
If chkcomb(allcount) = 1 Then
    If lblcomb.Caption = "" Then ' if no number already entered then put number in
    lblcomb.Caption = chkcomb(allcount).Caption
    Else: lblcomb.Caption = lblcomb.Caption & "+" & chkcomb(allcount).Caption ' else
    End If                                                     'put "+" infront of number
    If hisequation(allcount).Caption <> "" Then ' if there is equation chosen for combining
Rem Do the same procedure as cmddetail does when opening this window but this
Rem time plot all the graphs,chosen to combine on blue window
    main.ScriptControl1.Reset
    main.ScriptControl1.language = "VBscript"
        If eqhistorytype(allcount) = True Then ' if Y(x) = true
        main.ScriptControl1.ExecuteStatement ("dim x")
        main.ScriptControl1.ExecuteStatement ("x = " & -10)
        hiscomb.CurrentX = -10
        hiscomb.CurrentY = main.ScriptControl1.Eval(equationhistory(allcount))
        For X = xmin To xmax Step 0.01
        Y = main.ScriptControl1.Eval(equationhistory(allcount))
            On Error Resume Next
            If Y >= ymin And Y <= ymax Then
            hiscomb.Line (hiscomb.CurrentX, _
            hiscomb.CurrentY)-(X, Y)
            hiscomb.CurrentX = X
            hiscomb.CurrentY = Y
            End If
        main.ScriptControl1.ExecuteStatement ("x = x + " & 0.01)
        Next X
        Else ' if X(y)=true
        main.ScriptControl1.ExecuteStatement ("dim y")
        main.ScriptControl1.ExecuteStatement ("y = " & -10)
        hiscomb.CurrentY = -10
        hiscomb.CurrentX = main.ScriptControl1.Eval(equationhistory(allcount))
        For Y = ymin To ymax Step 0.01
        X = main.ScriptControl1.Eval(equationhistory(allcount))
            hiscomb.Line (hiscomb.CurrentX, _
            hiscomb.CurrentY)-(X, Y)
            hiscomb.CurrentX = X
            hiscomb.CurrentY = Y
        main.ScriptControl1.ExecuteStatement ("y = y + " & 0.01)
        Next Y
        End If
    End If
End If
Next
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Rem on load add some instructions to the text box
Private Sub Form_Load()
Select Case langval
Case 1
Text2.Text = " In this window you may see all the past equations in detail. You may also combine equations"
Text2.Text = Text2.Text & "by clicking on the number and choosing those you want to see on blue screen."
Text2.Text = Text2.Text & "For every screen x=-10 to 10 and y=-10 to 10."
Case 2
Text2.Text = "In dit venster kan je alle verleden vergelijkingen in detail bekijken.Je kan ook"
Text2.Text = Text2.Text & "vergelijkingen samenvoegen door op de nummer te klikken die je in he"
Text2.Text = Text2.Text & "blauwe scherm wil zien.  Voor elk venster x=10 tot 10 en y = 10 tot 10."
Case 3
Text2.Text = "Dans ce fenêtre tu peux voir tous les équations que tu as déjà fait.  Tu peux"
Text2.Text = Text2.Text & "aussi combiner des équations par selectioner le numéro que tu veux"
Text2.Text = Text2.Text & "voir dans la fenêtre bleue.  Pour chaque fenêtre x=10 à 10 et y=10 à 10."
End Select
End Sub

