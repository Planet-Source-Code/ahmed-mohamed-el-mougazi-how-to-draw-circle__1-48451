Attribute VB_Name = "Module1"
'This Code is Done By Ahmed Mohamed El-Mougazi
Option Explicit
Dim Mx As Single, My As Single, Mx1 As Single, My1 As Single, Mshape As Shape
Sub Draw(ByVal X As Single, ByVal Y As Single, ByVal X1 As Single, ByVal Y1 As Single, Shape1 As Shape)
Attribute Draw.VB_Description = "Draw A shape as A circle"
On Error GoTo 100
Shape1.Visible = True
If X1 > X Then
Shape1.Width = X1 - X
Shape1.Left = X
Else
Shape1.Width = X - X1
Shape1.Left = X1
End If
If Y1 > Y Then
Shape1.Height = Y1 - Y
Shape1.Top = Y
Else
Shape1.Height = Y - Y1
Shape1.Top = Y1
End If
Mx = X
My = Y
Mx1 = X1
My1 = Y1
Set Mshape = Shape1
Exit Sub
100
MsgBox Err.Description, vbCritical, "Error Number: " & Err.Number
End Sub
Sub Draw_Circle(Optional ByVal Color As Long = 0)
Attribute Draw_Circle.VB_Description = "Draw Circle With Dimensions Of Shape in Function Draw"
Dim Aspect As Single, X As Single, Y As Single, Width&
With Form1
Aspect = .Shape1.Height / .Shape1.Width
X = .Shape1.Left + (.Shape1.Width / 2)
Y = .Shape1.Top + (.Shape1.Height / 2)
If .Shape1.Width > .Shape1.Height Then
Width = .Shape1.Width / 2
Else
Width = .Shape1.Height / 2
End If
Form1.Circle (X, Y), Width, Color, , , Aspect
.Shape1.Visible = False
End With
End Sub
