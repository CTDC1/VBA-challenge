VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Alpha_Testing()

' Defining all our variables that we will use throughout the loop
Dim Ticker As String
Dim OpenChange As Double
Dim CloseChange As Double
Dim TotalChange As Double
Dim PercentageChange As Double
Dim SummaryRow As Double
Dim TotalVolume As LongLong
Dim LastRow As Double
Dim GreatestIncrease As String
Dim GreatestDecrease As String
Dim GreatestVolume As LongLong
Dim GreatestTicker1 As String
Dim GreatestTicker2 As String
Dim GreatestTicker3 As String
Dim ws As Worksheet

' Loop through all the worksheets and seeting ws to be shorthand for thisworkbook.worksheets
 For Each ws In ThisWorkbook.Worksheets
 
 ' Using a common forumla to calcualte the last row in the worksheet, this way we do not need to have the exact number everytime
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
 
 'Setting our counter variables to zero at the start of each worksheet
    TotalVolume = 0
    OpenChange = 0
    GreatestVolume = 0
    GreatestIncrease = 0
    GreatestDecrease = 0
 

 ' before we start collecting the data, we add in our titles for the summary and for the greatests.
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Total Change"
    ws.Cells(1, 11) = "Percentage Change"
    ws.Cells(1, 12) = "Total Volume"
    ws.Range("o1").Value = "Ticker"
    ws.Range("p1").Value = "Value"
    ws.Range("n2").Value = "Greatest % Increase"
    ws.Range("n3").Value = "Greatest % Decrease"
    ws.Range("n4").Value = "Greatest Volume"
 
' Define SummaryRow, Summaryrow is our code in the rows which the results of the loop will be added
 SummaryRow = 2
 
 ' Starting the loop, start at 2 because we have titles in 1, we do not need a exact number because we are using the lastrow variable that we set earlier
    For i = 2 To LastRow

' if statement comparing the current cell against the stored value of GreatestVolume and Selecting the Ticker Symbol as well
        If ws.Cells(i, "L").Value >= GreatestVolume Then
            GreatestVolume = ws.Cells(i, "L").Value
            GreatestTicker3 = ws.Cells(i, "I").Value
    
        End If
    
    
' if statement setting the open value to the first open price for each ticker, we want to make sure we have an if statement here because we don't want the openchange to change for each new row
        If OpenChange = 0 Then
            OpenChange = ws.Cells(i, 3).Value
        
        End If
    
' Adding the rows volume to the totalvolume counter, this does not need an if statement because we want the value from reach row
    TotalVolume = TotalVolume + ws.Cells(i, 7)

' If the next Ticker in Column A is a different than the next Ticker, start this if statement, This is now starting our calculations from the summary table

         If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
              Ticker = ws.Cells(i, 1).Value
              CloseChange = ws.Cells(i, 6).Value
    
' Calculating the Total Change, Percentage Change, and Total Volume
            TotalChange = CloseChange - OpenChange
            PercentageChange = (TotalChange / OpenChange)

' if statement comparing the current cell against the stored value of GreatestIncrease, we have this after the PercentageChange Calucation so it is accurate
                     If PercentageChange >= GreatestIncrease Then
                              GreatestIncrease = PercentageChange
                              GreatestTicker1 = ws.Cells(i, "I").Value
            
                    End If
    
' if statement comparing the current cell against the stored value of GreatestDecrease, we have this after the PercentageChange Calucation so it is accurate
                     If PercentageChange >= GreatestIncrease Then
                   If PercentageChange <= GreatestDecrease Then
                             GreatestDecrease = PercentageChange
                             GreatestTicker2 = ws.Cells(i, "I").Value
            
                    End If
         
'adding our variables to the summary row
    
              ws.Cells(SummaryRow, 9).Value = Ticker
              ws.Cells(SummaryRow, 10).Value = TotalChange
              ws.Cells(SummaryRow, 11).Value = PercentageChange
              ws.Cells(SummaryRow, 12).Value = TotalVolume


'create an conditional statement for the Total Change to be green or red, depending on whether its postive or negative
             If ws.Cells(SummaryRow, 10).Value >= 0 Then
                  ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            Else
                  ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
    
            End If
            
            
' Increase the count for summaryrow table row and reset our TotalVolume and OpenChange to zero so they can work for the next ticker
    SummaryRow = SummaryRow + 1
    TotalVolume = 0
    OpenChange = 0
    End If
    
        
        
        Next i
        'resetting summaryrow for the next worksheet
        SummaryRow = 2
    
 'once the loop is complete on the worksheet, I am setting the format for Percentage Change, Greatest Increase, and Greatest Decrease
ws.Range("p2").NumberFormat = "00.00%"
ws.Range("p3").NumberFormat = "00.00%"
ws.Range("K:K").NumberFormat = "00.00%"

'since I have the stored value for all these variables, I want to place them in the individual cells before setting the value back to 0 for the next worksheet
ws.Range("o2").Value = GreatestTicker1
ws.Range("p2").Value = GreatestIncrease

ws.Range("o3").Value = GreatestTicker2
ws.Range("p3").Value = GreatestDecrease

ws.Range("o4").Value = GreatestTicker3
ws.Range("p4").Value = GreatestVolume

'moving through the years
Next ws

    
End Sub




