Attribute VB_Name = "Module1"
'Script to loop through each year of stock data and calculate total volume by ticker symbol

Sub Ticker_Total()

  ' Set Active Worksheets
    Dim ws As Worksheet
    Dim Sheet1_ws As Worksheet
    
    Set Sheet1_ws = ActiveSheet
    
  ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
        ' Create a Variable to Hold File Name
        Dim WorksheetName As String
       
        ' Set an initial variable and place holder for holding the Ticker Symbol
        Dim Ticker_Symbol As String
        Cells(1, 9).Value = "Ticker Symbol"
        
        ' Set an initial variable and place holder for holding the Yearly Change
        Dim Yearly_Change As Long
        Yearly_Change = 0
        Cells(1, 10).Value = "Yearly Change"
        
        ' Set an initial variable for Opening Value
        Dim Opening_Value As Long
        Opening_Value = 0
        
        ' Set an initial variable for Closing Value
        Dim Closing_Value As Long
        Closing_Value = 0
        
        'Set an initial variable and place holder for holding the Percent Change
        Dim Percent_Change As Long
        Percent_Change = 0
        Cells(1, 11).Value = "Percent Change"
        
        ' Set an initial variable and place holder for holding the Ticker Total Volume
        Dim Ticker_Total As Double
        Ticker_Total = 0
        Cells(1, 12).Value = "Total Stock Volume"

        ' Keep track of the location for each Ticker Symbol in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        'Find last row in worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through all Ticker Symbols
        For i = 2 To LastRow

            ' Check if we are still within the same Ticker Symbol, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                ' Set the Ticker Symbol
                Ticker_Symbol = Cells(i, 1).Value

                ' Add to the Ticker Total
                Ticker_Total = Ticker_Total + Cells(i, 7).Value
                
                'Add Opening Value
                Opening_Value = Opening_Value + Cells(i, 3).Value
                
                'Add Closing Value
                Closing_Value = Closing_Value + Cells(i, 6).Value
                
                'Calculate the Yearly Change
                Yearly_Change = Closing_Value - Opening_Value
                
                'Calculate the Percent Change
                'Percent_Change = Yearly_Change / Ticker_Total

                ' Print the Ticker Symbol in the Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker_Symbol
                
                ' Print the Yearly Change in the Summary Table
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                ' Print the Percent Change in the Summary Table
                Range("K" & Summary_Table_Row).Value = Percent_Change

                ' Print the Ticker Total in the Summary Table
                Range("L" & Summary_Table_Row).Value = Ticker_Total

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the Ticker Total
                Ticker_Total = 0

            ' If the cell immediately following a row is the same brand...
            Else

                ' Add to the Ticker Total
                Ticker_Total = Ticker_Total + Cells(i, 7).Value

            End If

        Next i
        
    Next ws
    
End Sub
