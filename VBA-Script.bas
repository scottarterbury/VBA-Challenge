Attribute VB_Name = "Module1"
Sub Stocks():

 Dim ws As Worksheet
 
 For Each ws In Worksheets
  ws.Activate

    'space for variables location in summary table
    
    Dim Summary_Table As Integer
    Summary_Table = 2
    
    'show headings for summary table
    
    Cells(1, 12).Value = "Tickers"
    Cells(1, 13).Value = "Price Change"
    Cells(1, 14).Value = "Percentage Change"
    Cells(1, 15).Value = "Total Volume"
    
    
    'create the variable for all data
    
    Dim Last_Row As Long
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'create the variable for ticker
    
    Dim Ticker As String
    
    'create initial variable for total volume
    
    Dim Volume As Variant
    Volume = 0
    
    'create variable for price change
    
    Dim Price_Change As Double
    Price_Change = 0
    
    'create variable for opening value
    
    Dim Opening As Double
    Opening = 0
    
    'create variable for closing value
    
    Dim Closing As Double
    Closing = 0
    
    'create variable for percentage change
    
    Dim Percent_Change As Double
    Percent_Change = 0
    
    'create variables for table and max percent
           
    Dim Max_Percent As Double
    Max_Percent = 0
    
    'create variable for max ticker
    
    Dim Max_Ticker As String
    
    'create variable for min percent
    
    Dim Min_Percent As Double
    Min_Percent = 0
    
    'create variable for min ticker
    
    Dim Min_Ticker As String
    
    'create variable for max vol
    
    Dim Max_Volume As Variant
    Max_Volume = 0
    
    'create variable for max vol ticker
    
    Dim Max_Vol_Ticker As String


    'create the opening price
    
    Opening = Cells(2, 3).Value
    
    
    'loop for the ticker
    For i = 2 To Last_Row

        'check ticker is the same
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'set the ticker
            
            Ticker = Cells(i, 1).Value
            
            'add the total stock volume
            
            Volume = Volume + Cells(i, 7).Value
            
            'set closing price
            
            Closing = Cells(i, 6).Value

            'determine price change
            
            Price_Change = Closing - Opening
            
            'determine percentage change
            
                If Opening <> 0 Then
                    Percent_Change = (Price_Change / Opening)
                End If
                    
                
            'print ticker into sum table
            
            Range("L" & Summary_Table).Value = Ticker
            
            'print volume into sum table
            
            Range("O" & Summary_Table).Value = Volume
            
            'print price change into sum table
            
            Range("M" & Summary_Table).Value = Price_Change
                
                'color changes
                
                If Price_Change > 0 Then
                    Range("M" & Summary_Table).Interior.ColorIndex = 4
                Else
                    Range("M" & Summary_Table).Interior.ColorIndex = 3
                End If
            
            
            'format of percentage change
            
            Range("N" & Summary_Table).NumberFormat = "0.00%"
            
            'print percentage change into sum table
            
            Range("N" & Summary_Table).Value = Percent_Change
            
            'add one to summary table
            
            Summary_Table = Summary_Table + 1
            
            
            'reset volume, opening, closing price change
            
            Price_Change = 0
            Opening = Cells(i + 1, 3).Value
            Closing = 0
                
                'set values for table and determine min max percent change
                
                If Percent_Change > Max_Percent Then
                    Max_Percent = Percent_Change
                    Max_Ticker = Ticker
                ElseIf Percent_Change < Min_Percent Then
                    Min_Percent = Percent_Change
                    Min_Ticker = Ticker
                End If
                
                'determine highest volume with ticker
                
                If Volume > Max_Volume Then
                    Max_Volume = Volume
                    Max_Vol_Ticker = Ticker
                End If
            
            
            'reset percent and volume change
            
            Volume = 0
            Percent_Change = 0
            
        'see if next row starts with the same
        
        Else
        
            'add volume
            
            Volume = Volume + Cells(i, 7).Value
        
        End If
    
    Next i
    
    'create space for min max percent change location in table
    
    Dim Advanced_Table As Integer
    Advanced_Table = 2
    
    'create headings for table
    
    Cells(1, 18).Value = "Ticker"
    Cells(1, 19).Value = "Value"
    Cells(2, 17).Value = "Greatest % Change"
    Cells(3, 17).Value = "Lowest % Change"
    Cells(4, 17).Value = "Greates Total Volume"
    
    
    'format of min max percent
    
    Range("S2").NumberFormat = "0.00%"
    Range("S3").NumberFormat = "0.00%"
    
    'print variables to table
    
    Range("R2").Value = Max_Ticker
    Range("S2").Value = Max_Percent
    Range("R3").Value = Min_Ticker
    Range("S3").Value = Min_Percent
    Range("R4").Value = Max_Vol_Ticker
    Range("S4").Value = Max_Volume

    'resize cells
    
    ws.Cells.EntireColumn.AutoFit
    
 Next ws
            
    
End Sub

