Attribute VB_Name = "Module1"

Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
    ' Set and initial variable for holding the ticker name
Dim TickerName As String
' Set and initial variable for holding the total valume per ticker
Dim TotalStuck As Double
TotalStuck = 0

' Keep track of the location for each Ticker in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Loop through all ticker names
For i = 2 To 797711

'Check if we are still within the same ticker, if it is not...
  If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

' Set the ticker name
  TickerName = Cells(I, 1).Value

' Add to the total stuck volume
  TotalStuck = TotalStuck + Cells(I, 7).Value
  
' Print the ticker names in the summary table
  Range("I" & Summary_Table_Row).Value = TickerName
  
' Print the the total stock volume to the summary table
  Range("J" & Summary_Table_Row).Value = TotalStuck
  
' Add one to the summary table row
  Summary_Table_Row = Summary_Table_Row + 1
  
' Reset the total stuck
  TotalStuck = 0
  
' If the cell immediately following a row is the same ticker
Else

  ' Add to the total stuck
  TotalStuck = TotalStuck + Cells(I, 7).Value
  
End If

Next I

End Sub


