Attribute VB_Name = "plotter"

Public accuracy As Currency
Rem this function is the actual plotter
Rem it receives equation as a string,evaluates it and plots on picture
Rem eqhistype is added to distinguish between fraph types
Public Function plot(equation As String, eqhistype As Boolean)
main.screen.Scale (xmin, ymax)-(xmax, ymin) ' set scale
main.ScriptControl1.language = "VBScript" ' set scripting language
If eqhistype = True Then ' if Y(x) chosen then
    main.ScriptControl1.ExecuteStatement ("Dim x") ' dec x
    main.ScriptControl1.ExecuteStatement ("x = " & xmin) ' set x to its min value
    main.screen.CurrentX = xmin ' in case if plotting with lines set current coordinates
    main.screen.CurrentY = main.ScriptControl1.Eval(equation) ' to its starting position
    For X = xmin To xmax Step accuracy ' step is updated
    Y = main.ScriptControl1.Eval(equation) ' calculate y for given X
    On Error Resume Next
    If drawline = False Then ' if drawing with points
    If Y >= ymin And Y <= ymax Then main.screen.PSet (X, Y) ' if points are not out
    ' of range then plot them
    On Error Resume Next
    Else ' if plotting with lines
    If Y >= ymin And Y <= ymax Then ' if not out of range
    main.screen.Line (main.screen.CurrentX, main.screen.CurrentY)-(X, Y) 'draw lines
    main.screen.CurrentX = X
    main.screen.CurrentY = Y
    End If
    End If
    main.ScriptControl1.ExecuteStatement ("x = x + " & accuracy) 'increase value of x
    Next
    
Else ' if graph type is X(y) then repeat a little differently
    main.ScriptControl1.ExecuteStatement ("Dim y") 'dec y
    main.ScriptControl1.ExecuteStatement ("y = " & ymin) ' set y to min
    main.screen.CurrentY = ymin ' in case if plotting with lines
    main.screen.CurrentX = main.ScriptControl1.Eval(equation)
    For Y = ymin To ymax Step accuracy
    X = main.ScriptControl1.Eval(equation)
    If drawline = False Then
    main.screen.PSet (X, Y)
    On Error Resume Next
    Else
    main.screen.Line (main.screen.CurrentX, main.screen.CurrentY)-(X, Y)
    main.screen.CurrentX = X
    main.screen.CurrentY = Y
    End If
    main.ScriptControl1.ExecuteStatement ("y = y + " & accuracy)
    Next
End If
End Function
' roundnum instead of round because round will round number like 0.005 to 0 and i want 0.01
Public Function Roundnum(ByVal rNumber As Double, ByVal intDecimals As Integer) As Double
   Dim factor As Double
   Dim Temp As Double
   factor = 10 ^ intDecimals
   Temp = rNumber * factor + 0.5
   Roundnum = Int(Temp) / factor
End Function

