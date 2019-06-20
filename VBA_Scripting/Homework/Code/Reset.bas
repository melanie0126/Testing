Attribute VB_Name = "Module1"
Sub reset()
For Each ws In Worksheets

ws.Range("I:N").Clear
Next ws
End Sub
