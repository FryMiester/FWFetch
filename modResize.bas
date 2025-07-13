Attribute VB_Name = "modResize"
Option Explicit

Private Type ControlProportions
    WidthProp As Single
    HeightProp As Single
    TopProp As Single
    LeftProp As Single
End Type
Dim ArrayOfProportions() As ControlProportions

Public Sub InitResizeArray(theForm As Form)
Dim i As Integer
ReDim ArrayOfProportions(0 To theForm.Controls.Count - 1)
On Error Resume Next
    For i = 0 To theForm.Controls.Count - 1
        With ArrayOfProportions(i)
            .WidthProp = theForm.Controls(i).Width / theForm.ScaleWidth
            .HeightProp = theForm.Controls(i).Height / theForm.ScaleHeight
            .LeftProp = theForm.Controls(i).Left / theForm.ScaleWidth
            .TopProp = theForm.Controls(i).Top / theForm.ScaleHeight
        End With
    Next i
End Sub

Public Sub ResizeControls(theForm As Form)
Dim i As Integer
On Error Resume Next
    For i = 0 To theForm.Controls.Count - 1
        With ArrayOfProportions(i)
            theForm.Controls(i).Move .LeftProp * theForm.ScaleWidth, _
                                    .TopProp * theForm.ScaleHeight, _
                                    .WidthProp * theForm.ScaleWidth, _
                                    .HeightProp * theForm.ScaleHeight
        End With
    Next i
End Sub
