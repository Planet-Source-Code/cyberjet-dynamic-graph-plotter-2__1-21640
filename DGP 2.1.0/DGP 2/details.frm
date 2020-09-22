VERSION 5.00
Begin VB.Form details 
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdok 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtcorrinput 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton optcall 
      Caption         =   "QUALITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton color 
      Caption         =   "COLOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lbltype 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colorchosen As Boolean
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
main.screen.ForeColor = main.screen.BackColor
plot corrfunc, corrtypefunc
If txtcorrinput.Text <> "" Then  ' if textbox contains equation then...
On Error GoTo errors
main.screen.DrawWidth = 1 ' width of the graph
main.ScriptControl1.language = "VBScript" ' set script language to VBScript
main.screen.ForeColor = main.lblhistory(histstack).BackColor
estring = LCase(txtcorrinput.Text) ' assign the contents of textbox to string estring
SqueezeSpaces estring ' call function SqueezeSpaces to delete any space between constants and variables
txtcorrinput.Text = estring  ' assign squeezed text back to equation text box
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
If corrtypefunc = True Then ' get the type of a function
plot txtcorrinput.Text, True  ' pass an equation to the plot function, with function type
Else: plot txtcorrinput.Text, False
End If
End If
corrfunc = txtcorrinput.Text
main.lblhistory(histstack).Caption = lbltype.Caption & txtcorrinput.Text
main.plotaxis
If txtcorrinput.Text = "" Then main.lblhistory(histstack).Visible = False
Exit Sub
errors: ' this will handle error by displaying what the error is
res& = MsgBox("Error #" & CStr(Err.Number) & " " & Err.Description + vbCrLf + cstring$, vbOKOnly)
On Error GoTo 0
Exit Sub
End Sub

Private Sub color_Click()
colorchosen = True
Dim Cl As OLE_COLOR
    On Error Resume Next
    main.CD1.Flags = 0 ' set flags to 0
    main.CD1.DialogTitle = "Choose Color for your next graph" 'caption
    main.CD1.ShowColor ' show th window
   main.lblhistory(histstack).BackColor = main.CD1.color   'set color
End Sub

Private Sub optcall_Click()
frmOptions.Show vbModal
End Sub
Sub form_load()
Me.Left = main.Frame1(5).Left
Me.Top = main.Frame1(5).Top
If corrtypefunc = True Then
lbltype.Caption = "X="
Else: lbltype.Caption = "Y="
End If
colorchosen = False
End Sub


