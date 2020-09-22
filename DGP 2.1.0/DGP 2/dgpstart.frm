VERSION 5.00
Begin VB.Form dgpstart 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2220
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   2220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   0
      Width           =   720
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   4080
         Top             =   6720
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   3360
         Top             =   6720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   960
         TabIndex        =   1
         Top             =   3720
         Width           =   12375
      End
   End
End
Attribute VB_Name = "dgpstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x, z, o, m, change As Long
Dim whitecolor, endanim As Boolean

Private Sub Form_Load()
Picture1.Width = screen.Width
Picture1.Height = screen.Height
End Sub

Private Sub Label1_Click()
main.Show
Unload Me
End Sub
Private Sub Label2_Click()
main.Show
Unload Me
End Sub

Private Sub Picture1_Click()
main.Show
Unload Me
End Sub
Sub changecolor()
Label1.ForeColor = RGB(x, 0, z)
End Sub
Sub changecolor1()
Label1.ForeColor = RGB(o, o, 0)
End Sub

Private Sub Timer1_Timer()
counterval
change = change + 5
If change < 255 Then
    If whitecolor = False Then
    clrchange
    Else: clrchangew
    End If
    ElseIf change < 505 Then
        If whitecolor = False Then
        clrchange1
        Else: clrchangew1
        End If
        Else
        change = 0: z = 0
        changecolor
        End If
End Sub
Sub position()
Label1.Left = (screen.Width / 2 - Label1.Width / 2)
Label1.Top = (screen.Height / 2 - Label1.Height / 2)
End Sub
Private Sub Timer2_Timer()
Y = Y + 5
counterval
End Sub
Sub counterval()
If Y < 510 Then
Label1.FontSize = 32
Label1.Caption = "Alien Skin Software"
position
whitecolor = False
endanim = False
ElseIf Y < 1000 Then
Label1.FontSize = 16
Label1.Caption = "presents"
position
whitecolor = True
endanim = False
ElseIf Y < 1480 Then
Label1.FontSize = 38
Label1.Caption = "Dynamic Graphics Plotter 2"
position
whitecolor = False
endanim = True
Timer3.Enabled = True
Else
Exit Sub
End If
End Sub
Sub clrchange()
If endanim = False Then
x = x + 1
changecolor
Else: m = m + 1
changecolor1
End If
End Sub
Sub clrchange1()
If endanim = False Then
x = x - 1
changecolor
Else: m = m - 1
changecolor1
End If
End Sub
Sub clrchangew()
x = x + 1
z = z + 1
changecolor
End Sub
Sub clrchangew1()
x = x - 1
z = z - 1
changecolor
End Sub
Private Sub Timer3_Timer()
If o < 255 Then
o = o + 4
Else
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Label1_Click
Exit Sub
End If
changecolor1
End Sub
