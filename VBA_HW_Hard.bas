Attribute VB_Name = "Hard"
Sub stockgreatesttotals()

' Run through all worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
    
' Define variables for totals
    Dim gvolume As Double
    Dim gpercentd As Double
    Dim gpercenti As Double
        
' Set initial variable for ticker
    Dim ticker As String
    
' Define last worksheet row
    Dim lrow As Long

' Find the last non-blank cell in column A(1)
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Find greatest total volume
        gvolume = Application.WorksheetFunction.Max(ws.Columns("J"))
        ws.Cells(4, 16).Value = gvolume
            
    ' Find greatest % increase
        gpercenti = Application.WorksheetFunction.Max(ws.Columns("L"))
        ws.Cells(2, 16).Value = gpercenti
           
    ' Find greatest % decrease
        gpercentd = Application.WorksheetFunction.Min(ws.Columns("L"))
        ws.Cells(3, 16).Value = gpercentd
        
        ' For loop for ticker labels
            For g = 2 To lrow
            
            ' Identify closing ticker label
                ticker = ws.Cells(g, 9).Value
        
            ' Find and print ticker for greatest percent increase
                If ws.Cells(g, 12).Value = gpercenti Then
                ws.Cells(2, 15).Value = ticker
                End If
                
            ' Find and print ticker for greatest percent decrease
                If ws.Cells(g, 12).Value = gpercentd Then
                ws.Cells(3, 15).Value = ticker
                End If
                
            ' Find and print ticker for greatest total volume
                If ws.Cells(g, 10).Value = gvolume Then
                ws.Cells(4, 15).Value = ticker
                End If
                                                  
            Next g

    Next ws

End Sub

