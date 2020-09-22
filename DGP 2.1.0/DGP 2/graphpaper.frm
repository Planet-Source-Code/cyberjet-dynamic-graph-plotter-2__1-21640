VERSION 5.00
Begin VB.Form gpaper1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create a graph paper"
   ClientHeight    =   3375
   ClientLeft      =   2985
   ClientTop       =   2640
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5160
   Begin VB.Frame fraPrint 
      Caption         =   "Printer properties"
      Height          =   975
      Left            =   2640
      TabIndex        =   20
      Top             =   1200
      Width           =   2415
      Begin VB.ListBox lstCopies 
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblcopies 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Copies"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraScale 
      Caption         =   "Scale"
      Height          =   975
      Left            =   2640
      TabIndex        =   17
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton optCentimeter 
         Caption         =   "Centimeters"
         Height          =   255
         Left            =   360
         MaskColor       =   &H0080FFFF&
         TabIndex        =   19
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optInch 
         Caption         =   "Inches"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame fraLines 
      Caption         =   "Design"
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   2415
      Begin VB.ListBox lstColor 
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.ListBox lstThick 
         Height          =   255
         ItemData        =   "graphpaper.frx":0000
         Left            =   1200
         List            =   "graphpaper.frx":0002
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblColor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblThickness 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Thickness"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame fraThickLines 
      Caption         =   "Lines per square"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
      Begin VB.ListBox lstWideV 
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox lstWideH 
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblThickV 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vertical"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblThivkH 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Horizonal"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraNumLines 
      Caption         =   "Number of Lines"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.ListBox lstLinesV 
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.ListBox lstLinesH 
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLinesV 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vertical"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLinesH 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Horizonal"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "gpaper1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This codes belongs to someone else, i cant remember the name, sorry
Dim Counter As Integer
Dim ScaleType As String
Option Explicit
Private Sub LoadListboxes()
    ' -- Load Values into listboxes and set default values

    For Counter = 1 To 50
        lstLinesH.AddItem Counter
        lstLinesV.AddItem Counter
        lstWideH.AddItem Counter
        lstWideV.AddItem Counter
        lstCopies.AddItem Counter
    Next
    For Counter = 1 To 5
    lstThick.AddItem Counter
    Next
    lstColor.AddItem "Black"
    'vbBlack &H0 Black
    lstColor.AddItem "Red"
    'vbRed   &HFF    Red
    lstColor.AddItem "Green"
    'vbGreen &HFF00  Green
    lstColor.AddItem "Yellow"
    'vbYellow    &HFFFF  Yellow
    lstColor.AddItem "Blue"
    'vbBlue  &HFF0000    Blue
    lstColor.AddItem "Magenta"
    'vbMagenta   &HFF00FF    Magenta
    lstColor.AddItem "Cyan"
    'vbCyan  &HFFFF00    Cyan
    'vbWhite &HFFFFFF    White
    lstLinesH.ListIndex = 9
    lstLinesV.ListIndex = 9
    lstWideH.ListIndex = 9
    lstWideV.ListIndex = 9
    lstThick.ListIndex = 0
    lstColor.ListIndex = 0
    lstCopies.ListIndex = 0
    optInch_Click   'select inch scale

End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
    ' --- Build and print Graph Paper
    On Error GoTo PrintFailed
    cmdprint.Enabled = False
    Dim ColorToPrint As Long
    Dim Thickness%, WideThickness%, LineH%, LineV%
    Dim WideH%, WideV%
    Dim PrnMaxH%, PrnMaxW%, PrnMinH%, PrnMinW%
    Dim Counter%, WideCount%, CopyCounter%, NumOfCopies%
    If Printer.ColorMode = 1 Then 'Black and white only
        ColorToPrint = vbBlack
    Else
        Select Case lstColor.ListIndex
            Case 0
                ColorToPrint = vbBlack
                'vbBlack &H0 Black
            Case 1
                ColorToPrint = vbRed
                'vbRed   &HFF    Red
            Case 2
                ColorToPrint = vbGreen
                'vbGreen &HFF00  Green
            Case 3
                ColorToPrint = vbYellow
                'vbYellow    &HFFFF  Yellow
            Case 4
                ColorToPrint = vbBlue
                'vbBlue  &HFF0000    Blue
            Case 5
                ColorToPrint = vbMagenta
                'vbMagenta   &HFF00FF    Magenta
            Case 6
                ColorToPrint = vbCyan
                'vbCyan  &HFFFF00    Cyan
            Case Else
                ColorToPrint = vbBlack
        End Select
    End If ' - finished with color
    ' --- set Line Thickness
    Thickness = lstThick.ListIndex
    If Thickness = 0 Then Thickness = 1
    ' -- Set the Wide Line Thickness
    WideThickness = (Thickness * 1.5) + 2
    ' - max size of drawing area of printer
    PrnMaxH = Printer.ScaleHeight
    PrnMaxW = Printer.ScaleWidth
    PrnMinH = 0 + WideThickness
    PrnMinW = 0 + WideThickness
    'LineH%, LineV%, WideH%, WideV%
    Select Case ScaleType
        Case "Inch"
            LineH = 1440 / lstLinesH.ListIndex
            LineV = 1440 / lstLinesV.ListIndex
        Case "Centimeter"
            LineH = 567 / lstLinesH.ListIndex
            LineV = 567 / lstLinesV.ListIndex
        Case Else
            LineH = 1440 / lstLinesH.ListIndex
            LineV = 1440 / lstLinesV.ListIndex
    End Select

    WideH = lstWideH.ListIndex
    WideV = lstWideV.ListIndex
    NumOfCopies = lstCopies.ListIndex
    If NumOfCopies = 0 Then NumOfCopies = 1
    For CopyCounter = 1 To NumOfCopies
            ' -- Horizonal Lines
            WideCount = 0
            For Counter = PrnMinH To PrnMaxH Step LineH
                If WideCount = 0 Then
                    WideCount = WideH - 1
                    Printer.DrawWidth = WideThickness
                Else
                    WideCount = WideCount - 1
                    Printer.DrawWidth = Thickness
                End If
                Printer.Line (PrnMinW, Counter)-(PrnMaxW, Counter), ColorToPrint
            Next
            ' -- Vertical Lines
            WideCount = 0
            For Counter = PrnMinW To PrnMaxW Step LineH
                If WideCount = 0 Then
                    WideCount = WideV - 1
                    Printer.DrawWidth = WideThickness
                Else
                    WideCount = WideCount - 1
                    Printer.DrawWidth = Thickness
                End If
                Printer.Line (Counter, PrnMinH)-(Counter, PrnMaxH), ColorToPrint
            Next
            'Printer.NewPage
            Printer.EndDoc  'finished printing
    Next 'copies
    'Printer.EndDoc  'finished printing
    cmdprint.Enabled = True
    Exit Sub
PrintFailed:
    MsgBox "There was a problem printing!"
    cmdprint.Enabled = True
    Exit Sub
End Sub


Private Sub Form_Load()
    LoadListboxes
End Sub

Private Sub mnuExit_Click()
    ' --- Menu item, exit program
    Unload Me
End Sub


Private Sub optCentimeter_Click()
    ' --- Select number of lines per centimeter
    fraNumLines.Caption = "Number of lines per Centimeter"
    ScaleType = "Centimeter"
End Sub

Private Sub optInch_Click()
    ' --- Select number of lines per inch
    fraNumLines.Caption = "Number of lines -- Inch"
    ScaleType = "Inch"
End Sub
