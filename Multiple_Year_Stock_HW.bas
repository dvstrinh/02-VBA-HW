Attribute VB_Name = "Module1"
Sub multiple_year_stock()
    'Create a variable for ticker
    Dim Ticker As String
    
    'Create a variable for volume
    Dim Volume_Total As Double
    Volume_Total = 0
    
    'Create variables for yearly change
    Dim Yearly_Change As Double
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Percent_Change As Double
    
    Year_Open = Cells(2, 3).Value
    
    'Location of compiled ticker and volume data
    Dim Compiled_Data_Row As Double
    Compiled_Data_Row = 2

    'Define last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through all traded volume for ticker
    For i = 2 To lastrow
        
        'Year close
        If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Year_Close = Cells(i, 6).Value
            
        End If
        
        
        
        'Check to see if we're within same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
             
            'Set ticker name
            Ticker = Cells(i, 1).Value
        
            'Add to volume total
            Volume_Total = Volume_Total + Cells(i, 7).Value
            
            'Yearly Change
            Yearly_Change = Year_Close - Year_Open
        
            'Percent Change
            If Year_Open = 0 Then
            Percent_Change = 0
            Else
            Percent_Change = Yearly_Change / Year_Open
            
            End If
            
            'Print ticker name
            Range("I" & Compiled_Data_Row).Value = Ticker
            
            'Print total volume traded for ticker
            Range("J" & Compiled_Data_Row).Value = Volume_Total
            
            'Print yearly change
            Range("K" & Compiled_Data_Row).Value = Yearly_Change
            
            'Print percent change
            Range("L" & Compiled_Data_Row).Value = Percent_Change
            Range("L" & Compiled_Data_Row).NumberFormat = "0.00%"
            
            'Add one to the compiled data row
            Compiled_Data_Row = Compiled_Data_Row + 1
            
            'Reset volume total
            Volume_Total = 0
            
            'Reset open
             Year_Open = Cells(i + 1, 3).Value
            
            'Reset yearly change
            Yearly_Change = 0
            
            'Reset percent change
            Percent_Change = 0
            
        'If cell following a row is the same ticker
        Else
            
            'Add to the volume total
            Volume_Total = Volume_Total + Cells(i, 7).Value
        
            
        End If
        
    Next i
    
    '---------------------------------------------------------------------------------
        
        'Last Row
        lastrow_2 = Cells(Rows.Count, 11).End(xlUp).Row
        
        'Conditional Formatting Loop
        For i = 2 To lastrow_2
        'Conditional Formatting
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.Color = RGB(0, 255, 0)
            
        Else
            Cells(i, 11).Interior.Color = RGB(255, 0, 0)
            
        End If
        
        Next i
        
    '---------------------------------------------------------------------------------
        
        'Greatest percent change, volume
        Dim G_Increase As Double
        Dim G_Increase_Ticker As String
        Dim G_Decrease As Double
        Dim G_Decrease_Ticker As String
        Dim G_Volume As Double
        Dim G_Volume_Ticker As String
        
        G_Increase = 0
        G_Decrease = 0
        G_Volume = 0
        
        
        For j = 2 To lastrow_2
        
            'Greatest percent increase
            If Cells(j, 12).Value > G_Increase Then
            G_Increase = Cells(j, 12).Value
            G_Increase_Ticker = Cells(j, 9).Value
            Range("P2") = G_Increase_Ticker
            Range("Q2") = G_Increase
            Range("Q2").NumberFormat = "0.00%"
        
            
            'Greatest percent decrease
            ElseIf Cells(j, 12).Value < G_Decrease Then
            G_Decrease = Cells(j, 12).Value
            G_Decrease_Ticker = Cells(j, 9).Value
            Range("P3") = G_Decrease_Ticker
            Range("Q3") = G_Decrease
            Range("Q3").NumberFormat = "0.00%"
            
            'Greatest volume
            ElseIf Cells(j, 10).Value > G_Volume Then
            G_Volume = Cells(j, 10).Value
            G_Volume_Ticker = Cells(j, 9).Value
            Range("P4") = G_Volume_Ticker
            Range("Q4") = G_Volume
            
            End If
        
    Next j

    
    
End Sub
