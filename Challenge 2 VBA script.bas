Attribute VB_Name = "Module1"
Sub StockReview()
    
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim Summarize_Data_Row As Long
    Dim GreatestIncrease As Double
    GreatestIncrease = 0
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    Dim GreatestVolume As Double
    GreatestVolume = 0
    Dim Greatest_Increased_Ticker As String
    Greatest_Increased_Ticker = ""
    Dim Greatest_Decreased_Ticker As String
    Greatest_Decreased_Ticker = ""
    Dim Greatest_Volume_Ticker As String
    Greatest_Volume_Ticker = ""
    
    
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
       
        Summarize_Data_Row = 2
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        
        Ticker = Cells(2, 1).Value
        OpenPrice = Cells(2, 3).Value
        TotalVolume = 0
        
       
        For i = 2 To LastRow
            If Cells(i + 1, 1).Value <> Ticker Then
            
                ClosePrice = Cells(i, 6).Value
                
            
                YearlyChange = ClosePrice - OpenPrice
                
               
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice)
                Else
                    PercentChange = 0
                End If
                
                ws.Cells(Summarize_Data_Row, 11).NumberFormat = "0.00%"
             
                Cells(Summarize_Data_Row, 9).Value = Ticker
                Cells(Summarize_Data_Row, 10).Value = YearlyChange
                Cells(Summarize_Data_Row, 11).Value = PercentChange
                Cells(Summarize_Data_Row, 12).Value = TotalVolume
                
               
                If YearlyChange > 0 Then
                    Cells(Summarize_Data_Row, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    Cells(Summarize_Data_Row, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                
                Summarize_Data_Row = Summarize_Data_Row + 1
                Ticker = Cells(i + 1, 1).Value
                OpenPrice = Cells(i + 1, 3).Value
                TotalVolume = 0
            End If
            
            
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
             Cells(1, 17).Value = "Ticker"
             Cells(1, 18).Value = "Value"
            
            
            If PercentChange > GreatestIncrease Then
                GreatestIncrease = PercentChange
                Greatest_Increased_Ticker = Ticker
            End If
            
            If PercentChange < GreatestDecrease Then
                GreatestDecrease = PercentChange
                Greatest_Decreased_Ticker = Ticker
            End If
            If TotalVolume > GreatestVolume Then
                GreatestVolume = TotalVolume
                Greatest_Volume_Ticker = Ticker
            End If
        Next i
               
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(2, 17).Value = Greatest_Increased_Ticker
        ws.Cells(3, 17).Value = Greatest_Decreased_Ticker
        ws.Cells(4, 17).Value = Greatest_Volume_Ticker
        ws.Cells(2, 18).Value = GreatestIncrease
        ws.Cells(3, 18).Value = GreatestDecrease
        ws.Cells(4, 18).Value = GreatestVolume
    Next ws
End Sub

