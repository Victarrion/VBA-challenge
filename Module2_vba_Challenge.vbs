Attribute VB_Name = "Module1"
Sub stock_picker()


'Module assignemnt text:
'Create a script that loops through all the stocks for one year and outputs the following information:
'   The ticker symbol
'   Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'   The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'   The total stock volume of the stock. The result should match the following image:
' Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".


'comment to explain my understanding of the data:
'So it appears column A shows the stock name, column b shows the date with the following columns describing the opening, max, min, and closing
' price. Lastly there is volume of stocks traded.

' For this module, the logic is that the code must go through each row and save the values of interest. Best way is to save an
' array saving the row values.
' If dates are ordered, the script is very straightfoward. I assume dates will not be ordered nor do I know begining date or end date. So i need to keep track
' of first date and last date for a stock. I also need to have a variable adding up the volumes.

' once i have this data per stock, i make another row scanner to pull the datas of interest for the last part.

' i also need to cycle through each worksheet


Dim WS As Worksheet                                 'counts number of worksheets
Dim row_counter As Double                               ' counter for number of rows in a worksheet
Dim bool As Boolean                                     ' logical switch used in row counter loop
                                                        'Sets the row counter to start at 2 to make space for heading

Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim greatest_name As String
Dim decrease_Name As String
Dim Greatest_total_volume As Double
Dim Greatest_total_volume_name As String

Dim n As Integer    'counts the number of tickers and helps organize the rows with results of interest

Dim Ticker_details(2) As Double
Dim Ticker_Name As String

'Ticker_details(0)=opening price
'Ticker_details(1)=closing price
'Ticker_details(2)=volume


For Each WS In ThisWorkbook.Worksheets


'Resets values when switching worksheet
n = 2
row_counter = 1
bool = True

Greatest_Increase = 0
Greatest_Decrease = 0
Greatest_total_volume = 0
Ticker_details(0) = 0
Ticker_details(1) = 0
Ticker_details(2) = 0



'Headers
WS.Cells(1, 10).Value = "Ticker"
WS.Cells(1, 11).Value = "Yearly change"
WS.Cells(1, 12).Value = "Percent Change"
WS.Cells(1, 13).Value = "Total Stock Volume"

WS.Cells(1, 16).Value = "Ticker"
WS.Cells(1, 17).Value = "Value"
WS.Cells(2, 15).Value = "Greatest % increase"
WS.Cells(3, 15).Value = "Greatest % decrease"
WS.Cells(4, 15).Value = "Greatest total Volume"

While bool = True
    If WS.Cells(row_counter, 1).Value <> "" Then                                                'counts until it reaches the last row
        row_counter = row_counter + 1
    Else
        bool = False                                                                            'breaks loop when it counter the first empty row
    End If
Wend
' This loop is super inefficient, there must be a better way.


'Now that I know what the last row is i can now do a for loop to move around the sheet.

Ticker_Name = WS.Cells(2, 1).Value                                                                 ' Gets first ticker
Ticker_details(0) = WS.Cells(2, 3).Value                                                           ' Establishes the first opening value
 
For i = 2 To row_counter
    
    If (WS.Cells(i, 1).Value <> WS.Cells(i + 1, 1).Value And WS.Cells(i, 2).Value > WS.Cells(i + 1, 2).Value) Then
    'if the date is less and the ticker change assume we have a new ticker
        
        Ticker_details(1) = WS.Cells(i, 6).Value                                                   'establishes closing price
        Ticker_details(2) = Ticker_details(2) + WS.Cells(i, 7).Value                               'last volume added
        
        WS.Cells(n, 10).Value = Ticker_Name                                                        'prints ticker name
        WS.Cells(n, 11).Value = Ticker_details(1) - Ticker_details(0)                              'prints ticker dollar change
        WS.Cells(n, 11).NumberFormat = "$0.00"                                                      'formats to two decimal points
        WS.Cells(n, 12).Value = (((Ticker_details(1) - Ticker_details(0)) / Ticker_details(0)))    'prints ticker percentage change
        WS.Cells(n, 12).NumberFormat = "0.00%"                                                     'formats to percentage
        WS.Cells(n, 13).Value = Ticker_details(2)                                                  'prints ticker total volume
        
        'color codes yearly change
        If WS.Cells(n, 11).Value >= 0 Then
            WS.Cells(n, 11).Interior.ColorIndex = 4
        ElseIf WS.Cells(n, 11).Value < 0 Then
            WS.Cells(n, 11).Interior.ColorIndex = 3
        End If
        
        ' compares values to see if the new row introduces the greatest gain/loss and saves the value
        If (WS.Cells(n, 12).Value > Greatest_Increase And WS.Cells(n, 12).Value > 0) Then
            Greatest_Increase = WS.Cells(n, 12).Value
            greatest_name = WS.Cells(n, 10).Value
        ElseIf (WS.Cells(n, 12).Value < Greatest_Decrease And WS.Cells(n, 12).Value < 0) Then
            Greatest_Decrease = WS.Cells(n, 12).Value
            decrease_Name = WS.Cells(n, 10).Value
            
        End If
        
        ' compares values to see if the new row introduces the greatest volume and saves the value
        ' not placed in the previous if block because the greatest gain/lost can also be greatest volume row. A decision tree would make it so one value wouldnt be saved.
        
        If WS.Cells(n, 13).Value > Greatest_total_volume Then
           Greatest_total_volume_name = WS.Cells(n, 10).Value
           Greatest_total_volume = WS.Cells(n, 13).Value
        End If
        
        
        
        Ticker_Name = WS.Cells(i + 1, 1).Value                                                     'establishes new ticker
        Ticker_details(0) = WS.Cells(i + 1, 3).Value                                               'saves new opening value
        Ticker_details(2) = 0                                                                      'resets volume total

        n = n + 1                                                                                  'moves ticker counter
        
    ElseIf (WS.Cells(i, 1).Value = WS.Cells(i + 1, 1).Value) Then
        Ticker_details(2) = Ticker_details(2) + WS.Cells(i, 7).Value                               'sums volume traded
    End If
Next i

'places saved values

WS.Cells(2, 16).Value = greatest_name
WS.Cells(2, 17).Value = Greatest_Increase
WS.Cells(2, 17).NumberFormat = "0.00%"

WS.Cells(3, 16).Value = decrease_Name
WS.Cells(3, 17).Value = Greatest_Decrease
WS.Cells(3, 17).NumberFormat = "0.00%"

WS.Cells(4, 16).Value = Greatest_total_volume_name
WS.Cells(4, 17).Value = Greatest_total_volume

Range("J:Q").EntireColumn.AutoFit

Next WS



End Sub


