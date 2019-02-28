Sub alphabet_testing()

    For Each ws In Worksheets

        ' Set an initial variable for holding the Ticker name
        Dim Ticker_Name As String

        ' Set an initial variable for holding the total Volume per Ticker
        Dim Ticker_Volume As Double
        Ticker_Volume = 0
  
        Dim Open_Price As Double
        Open_Price = 0
  
        Dim Close_Price As Double
        Close_Price = 0

        Dim Yearly_Change As Double
        Yearly_Change = 0
  
        Dim Percent_Change 'As Double
        Yearly_Change = 0
  
        'Percentage format
        'Range("P2").NumberFormat = "0.00%"
        'Range("P3").NumberFormat = "0.00%"
        'Range("P4").NumberFormat = "0000000000"
        ' Keep track of the location for each credit card brand in the summarytable
        Dim Summary_Table_Row As Integer
  
        Summary_Table_Row = 2
    
        Dim j As Integer
    
        j = 0
        'Dim lastrow As Integer
    
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through all Tickers
  
        'For i = 2 To 264
  
        For i = 2 To lastrow

                ' Check if we are still within the same Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Ticker Name
                Ticker_Name = ws.Cells(i, 1).Value

                ' Add to the Ticker_Volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
        
                Open_Price = ws.Cells(i - j, 3).Value
        
                Close_Price = ws.Cells(i, 6).Value
        
                Yearly_Change = Close_Price - Open_Price
                'Percent_Change = (Yearly_Change / Open_Price) * 100
                Percent_Change = FormatPercent(Yearly_Change / Open_Price)

                ' Print the Ticker Name in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Ticker_Name

                'Print the Brand Amount to the Summary Table
                ws.Range("N" & Summary_Table_Row).Value = Ticker_Volume
                'Print the Yearly Change to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Yearly_Change
                'Print the Percent Change to the Summary Table
                'ws.Range("M" & Summary_Table_Row).Value = Percent_Change
                 
                    If Open_Price = 0 & Close_Price <> 0 Then
                
                        ws.Range("M" & Summary_Table_Row).Value = 0
                    
                    Else
                        ws.Range("M" & Summary_Table_Row).Value = Percent_Change
                    End If
                
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                '  Reset the Ticker_Volume
    
                j = 0
        
                Ticker_Volume = 0
                'If the cell immediately following a row is the same brand...
            Else
                'Add to the Ticket_Volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
           
                j = j + 1
        
   
            End If

        Next i
        'Color Formatting Code
        lRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
        Dim counter As Integer
        
        For counter = 2 To lRow
    
            If (ws.Cells(counter, 12).Value > 0) Then
                ws.Cells(counter, 12).Interior.ColorIndex = 4
            Else
                ws.Cells(counter, 12).Interior.ColorIndex = 3
            End If
        Next counter
            'Greatest Percent Increase
            MRow = ws.Cells(Rows.Count, 13).End(xlUp).Row
            Dim counter1 As Integer
            Dim GPI
            GPI = 0
            Dim GPIT As String
  
    
        For counter1 = 2 To MRow

            If (GPI < ws.Cells(counter1, 13).Value) Then
               GPI = ws.Cells(counter1, 13).Value
               GPIT = ws.Cells(counter1, 11).Value
            Else
               'MsgBox ("no value")
            End If
        Next counter1
            'MsgBox (GPI)
                ws.Cells(2, 19).Value = FormatPercent(GPI)
                ws.Cells(2, 18).Value = GPIT
                
                'Greatest Percent Decrease
            NRow = ws.Cells(Rows.Count, 13).End(xlUp).Row
            Dim counter2 As Integer
            Dim GPD
            GPD = 0
            Dim GPDT As String
  
    
        For counter2 = 2 To NRow
    
            If (GPD > ws.Cells(counter2, 13).Value) Then
               GPD = ws.Cells(counter2, 13).Value
               GPDT = ws.Cells(counter2, 11).Value
               
            Else
               'MsgBox ("no value")
            End If
        Next counter2
            'MsgBox (GPD)
            ws.Cells(3, 19).Value = FormatPercent(GPD)
            ws.Cells(3, 18).Value = GPDT
  
            'Greatest Volume Increase
            NRow = ws.Cells(Rows.Count, 14).End(xlUp).Row
            Dim counter3 As Integer
            Dim GTV As Double
            GTV = 0
            Dim GTVT As String
  
        For counter3 = 2 To NRow
    
            If (GTV < ws.Cells(counter3, 14).Value) Then
                GTV = ws.Cells(counter3, 14).Value
                GTVT = ws.Cells(counter3, 11).Value
               
            Else
               'MsgBox ("no value")
            End If
        
        Next counter3
            ws.Cells(4, 19).Value = GTV
            ws.Cells(4, 18).Value = GTVT
        'MsgBox (GTV)
    Next ws
  
End Sub
  



