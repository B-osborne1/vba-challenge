VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Sub StockAnalysis():

    'Set code to run on all sheets  **Note 1 (Excel Destination.Youtube)
Dim a As Integer
a = Application.Worksheets.Count


    For J = 1 To a
    Worksheets(J).Activate
        
    'Start code
    
        'Label columns
        
            'Initial
            Range("K1") = "Ticker"
            Range("L1") = "Yearly Change"
            Range("M1") = "Yearly% change"
            Range("N1") = "Volume total"
            
            'Bonus
            Range("T1") = "Ticker"
            Range("U1") = "Value"
            Range("S2") = "Greatest% gained"
            Range("S3") = "Greatest% lost"
            Range("S4") = "Highest volume"
            
            
        'Set variables
            Dim Ticker As String
            
            'Change in price variables
            Dim StartPr As Double
            Dim EndPr As Double
            Dim ChangePr As Double
            
            'Change in percentage
            Dim ChangePe As Double
            
            'Change in volume
            Dim Volume As Double
            
            'New row per ticker
            Dim Row As Integer
        
            'Bonus variables
            Dim PerG As Double
            Dim PerL As Double
            Dim VolM As Double
            
        
        'Set values for numbered variables
        StartPr = Cells(2, 3).Value
        EndPr = 0
        ChangePr = 0
        ChangePe = 0
        Volume = 0
        Row = 2
        PerG = 0
        PerL = 0
        VolM = 0
        
        
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        'Set iterations
        For I = 2 To LastRow
        
        
            'Check for differences
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
            
                'Looking for:
                Ticker = Cells(I, 1)
                EndPr = Cells(I, 6).Value
                
                'Calculate
                ChangePr = EndPr - StartPr
                ChangePe = ((EndPr - StartPr) / StartPr) * 100
                Volume = Volume + Cells(I, 7).Value
                
                'Put it where
                Cells(Row, 11) = Ticker
                Cells(Row, 12) = ChangePr
                Cells(Row, 13) = ChangePe
                Cells(Row, 14) = Volume
                
                'Shift and reset
                Row = Row + 1
                Volume = 0
                '(i+1,3) should be the first <open> value of new ticker
                StartPr = Cells(I + 1, 3).Value
            EndPr = 0
            
            
            'Calculate total volume per
            Else: Volume = Volume + Cells(I, 7)
            
        End If
    Next I
            
        'Conditional formatiing
            
        For I = 2 To LastRow
        
            If Cells(I, 12).Value < 0 Then
                'Red for values in negatives
                Cells(I, 12).Interior.ColorIndex = 3
                Cells(I, 13).Interior.ColorIndex = 3
                Range("M" & I) = WorksheetFunction.Round(Range("M" & I), 2)
                
                'Green for values in positives
            ElseIf Cells(I, 12).Value > 0 Then
                Cells(I, 12).Interior.ColorIndex = 4
                Cells(I, 13).Interior.ColorIndex = 4
                Range("M" & I) = WorksheetFunction.Round(Range("M" & I), 2)
            
            ElseIf Cells(I, 12).Value = 0 Then
                Cells(I, 12).Interior.ColorIndex = 0
                Cells(I, 13).Interior.ColorIndex = 0
                
            End If
        
    Next I
            
         'Finding greatest% plus values
         
            For I = 2 To LastRow
            
                If Cells(I, 13) > PerG Then
                PerG = Cells(I, 13).Value
                Range("U2") = PerG
                Range("T2") = Cells(I, 11).Value
                
                Range("U2") = WorksheetFunction.Round(Range("U2"), 2)
                End If
            
            Next I
            
            
            'Finding greatest% loss values
         
            For I = 2 To LastRow
            
                If Cells(I, 13) < PerL Then
                PerL = Cells(I, 13).Value
                Range("U3") = PerL
                Range("T3") = Cells(I, 11).Value
                
                Range("U3") = WorksheetFunction.Round(Range("U3"), 2)
                End If
            
            Next I
            
            'Finding greatest volume values
         
            For I = 2 To LastRow
            
                If Cells(I, 14) > VolM Then
                VolM = Cells(I, 14).Value
                Range("U4") = VolM
                Range("T4") = Cells(I, 11).Value
                End If
        
            Next I
     
     
     'Clean up
     
        'Autofit all rows and columns  **Note 2 (Puneet - Excelchamps)
            ActiveSheet.UsedRange.EntireColumn.AutoFit
            ActiveSheet.UsedRange.EntireRow.AutoFit
        
        

    Next J
        
    
    
End Sub


