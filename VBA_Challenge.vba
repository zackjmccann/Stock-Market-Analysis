Sub VBAChallenge()
    'Provide Headers for Each Column
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percentage Change"
        Range("L1").Value = "Total Stock Volume"
    
    'Declare the Variable for the Script
        Dim ticker_name As Long
        Dim next_comp As Long
        Dim all_comps As Long
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim stock_volume As Double

        all_comps = Cells(Rows.Count, 1).End(xlUp).Row
        next_comp = 1
        stock_volume = 0

    'Loop through the first column and pull the information needed for each company
        For ticker_name = 2 To all_comps
            
            'If the previous company is different from the current...
                If Cells(ticker_name - 1, 1).Value <> Cells(ticker_name, 1).Value Then
                    
                    'Hold the place for the summary columns
                    next_comp = next_comp + 1
                    
                    'Print the company's ticker symbol
                    Cells(next_comp, 9).Value = Cells(ticker_name, 1).Value
                    
                    'Note the opening price
                    open_price = Cells(ticker_name, 3).Value
                    
                    'Reset the stock volume total to 0
                    stock_volume = 0
                
            'If the previous company is the same as the current...
                Else
                    
                    'Note the closing price
                    close_price = Cells(ticker_name, 6).Value
                    
                    'Calculate the yearly change for the company
                    yearly_change = close_price - open_price
                    
                    'Print the yearly change in the appropriate column
                    Cells(next_comp, 10).Value = yearly_change
                    
                    'Print the cumulative total stock volume in the appropriate column
                    stock_volume = stock_volume + Cells(ticker_name, 7).Value
                    Cells(next_comp, 12).Value = stock_volume
                    
                End If
            
            'Color code the yearly change column
            
                If yearly_change >= 0 Then
                    'Print negative numbers as red
                    Cells(next_comp, 10).Interior.ColorIndex = 4
                Else
                    'Print positve numbers and 0 as green
                    Cells(next_comp, 10).Interior.ColorIndex = 3
                End If
                
            'Calculate the percent change
                If (open_price + close_price) = 0 Then
                    percent_change = 0
                Else
                    percent_change = (yearly_change / ((open_price + close_price) / 2) * 100)
                End If
            'Print the percent change in the appropriate column
                Cells(next_comp, 11).Value = percent_change
            
        Next ticker_name
        
    'CHALLENGE SECTION

        'Labels for Challenge
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
        
        'Declare Variables
            Dim challenge_values As Double
            Dim greatest_increase As Double
            Dim greatest_decrease As Double
            Dim greatest_volume As Double
            
            greatest_increase = 0
            greatest_decrease = 0
            greatest_volume = 0
            
            For challenge_value = 2 To all_comps
            'Greatest Increase
                If Cells(challenge_value, 11).Value > greatest_increase Then
                    'Print the value
                    Range("P2").Value = Cells(challenge_value, 9).Value
                    Range("Q2").Value = Cells(challenge_value, 11).Value
                    greatest_increase = Cells(challenge_value, 11).Value
                End If
                
             'Greatest Decrease
                If Cells(challenge_value, 11).Value < greatest_decrease Then
                    'Print the value
                    Range("P3").Value = Cells(challenge_value, 9).Value
                    Range("Q3").Value = Cells(challenge_value, 11).Value
                    greatest_decrease = Cells(challenge_value, 11).Value
                End If
                
            'Greatest Volume
                If Cells(challenge_value, 12).Value > greatest_volume Then
                    'Print the value
                    Range("P4").Value = Cells(challenge_value, 9).Value
                    Range("Q4").Value = Cells(challenge_value, 12).Value
                    greatest_volume = Cells(challenge_value, 12).Value
                End If
                
            Next challenge_value
        
End Sub