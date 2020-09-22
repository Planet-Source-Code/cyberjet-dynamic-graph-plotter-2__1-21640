Attribute VB_Name = "scriptcontrol"

Public estring As String
Public Function NumOccStr(inval$, c$)
'Find Number of occurences of character c$ in inval$
NumOccStr = 0
p1 = 1
Do
  p2 = InStr(p1, inval$, c$)
  If p2 <> 0 Then NumOccStr = NumOccStr + 1 Else Exit Function
  p1 = p2 + 1
Loop
End Function

Public Sub FindMatchingClosingBracket(inval$, pin, pout)
'pin is the position of an (
'pout is the position of the matching )
pob = InStr(pin + 1, inval$, "(")
If pob = 0 Then  '() no intermediate brackets
   pout = InStr(pin + 1, inval$, ")")
   Exit Sub
Else  '( @ pob before )
   nopbr = 0: nocbr = 0
   For k = pin To Len(inval$)
      c$ = Mid$(inval$, k, 1)
      If c$ = "(" Then nopbr = nopbr + 1
      If c$ = ")" Then nocbr = nocbr + 1
      If nopbr = nocbr Then
         pout = k
         Exit Sub
      End If
   Next k
End If
End Sub
Public Sub SqueezeSpaces(inval$)
'Squeeze out all spaces, trim & remove any leading +
inval$ = Trim$(inval$)
pp = InStr(1, inval$, "+")
If pp = 1 Then inval$ = Mid$(inval$, 2)
Do
  ps = InStr(1, inval$, " ")
  If ps = 0 Then Exit Do
  inval$ = Left(inval$, ps - 1) + Mid$(inval$, ps + 1)
Loop
End Sub

