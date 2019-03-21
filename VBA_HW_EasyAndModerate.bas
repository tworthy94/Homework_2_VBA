Attribute VB_Name = "EasyAndModerate"
' Easy portion

Sub stocktotal()

' Run through all worksheets
Dim ws As Worksheet
For Each ws In Worksheets

' Set initial variable for ticker
Dim closingticker As String

' Set initial variable for holding total per ticker
Dim volume_total As Double
volume_total = 0

' Track total location in summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
    
' Define last worksheet row
Dim lrow As Long

' Find the last non-blank cell in column A(1)
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Set variable to store opening price in
Dim j As Long
    j = 2

' Loop for each year
    For I = 2 To lrow
            
    ' Determine when ticker changes
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
    ' Identify closing ticker label
        closingticker = ws.Cells(I, 1).Value
        
    ' Define variables for open price and closing price
        Dim openvalue As Double
        openvalue = ws.Cells(j, 3).Value
        Dim closingvalue As Double
        closingvalue = ws.Cells(I, 6).Value
        Dim valuechange As Double
        valuechange = closingvalue - openvalue
        Dim percentchange As Double
              
    ' Add the ticker volume to volume total
        volume_total = volume_total + ws.Cells(I, 7).Value
            
    ' Ignore 0 in opening value
        If openvalue <> 0 Then
        
        'Calculate percent change
            percentchange = valuechange / openvalue
            
        Else
        percentchange = 0
            
        End If
            
        ' Round percent change
            percentchange = Round(percentchange, 2)
            
        ' Print ticker in summary table
            ws.Range("I" & Summary_Table_Row).Value = closingticker

        ' Print volume total to summary table
            ws.Range("J" & Summary_Table_Row).Value = volume_total
            
        ' Print value change total in summary table
            ws.Range("K" & Summary_Table_Row).Value = valuechange

        ' Print percent change to summary table
            ws.Range("L" & Summary_Table_Row).Value = percentchange

        ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the totals
            volume_total = 0
            
        ' Advance J
            j = I + 1
            
        ' If the cell following has the same ticker
            
            Else
            
            ' Add the volume and value change totals
                volume_total = volume_total + ws.Cells(I, 7).Value
                
            End If
        
        Next I
    
    Next ws
    
End Sub
