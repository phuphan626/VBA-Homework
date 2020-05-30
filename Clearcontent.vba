Attribute VB_Name = "Module2"
Sub ClearContent()
Dim ws As Worksheet
For Each ws In Worksheets
    Range("I1:Q1000").Delete
Next ws

End Sub

