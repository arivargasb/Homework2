

Sub stocks()

 
'**********************************************
' 1. Define variables and table's titles
'**********************************************

' Define variables
    Dim change As Double
    change = 0
    Dim change_per As Double
    change_per = 0
    Dim price_initial As Double
    price_initial = 0
    Dim price_final As Double
    price_final = 0
    Dim table As String
    Dim table_row As Integer
    table_row = 2


'**********************************************
' 2. Stocks 2014
'**********************************************
'Activate sheet
    Worksheets("2014").Activate
    
    ' Table's titles
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly change"
    Cells(1, 11).Value = "Percent change"
    Cells(1, 12).Value = "Total stock volume"
   
' Define last row and column
    lrow = Cells(Rows.Count, 2).End(xlUp).Row

       For i = 2 To lrow
' If last observation of the same stock
           If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
' Add volume of each day
               volume = volume + Cells(i, 7).Value
' Extract info if first or last day of the year
                If Cells(i, 2).Value = 20140101 Then
                    price_initial = Cells(i, 3).Value
                    volume_initial = Cells(i, 7).Value
                ElseIf Cells(i, 2).Value = 20141231 Then
                    price_final = Cells(i, 6).Value
                End If
                stock = Cells(i, 1).Value
' Calculate stock values
                change = price_final - price_initial
                If price_initial = 0 Then
                    change_per = 0
                Else
                    change_per = (price_final / price_initial) - 1
                End If
' Print stock values
                Worksheets("2014").Cells(table_row, 9).Value = stock
                Worksheets("2014").Cells(table_row, 10).Value = change
                Worksheets("2014").Cells(table_row, 11).Value = change_per
                Worksheets("2014").Cells(table_row, 12).Value = volume
' Add one to table row
                table_row = table_row + 1
' Reset values
                stock = 0
                change = 0
                change_per = 0
                volume = 0
            Else
' If not last observation of the same stock --> add volume for each day
               volume = volume + Cells(i, 7).Value
' And only extract info if first or last day of the year
                If Cells(i, 2).Value = 20140101 Then
                    price_initial = Cells(i, 3).Value
                ElseIf Cells(i, 2).Value = 20141231 Then
                     price_final = Cells(i, 6).Value
                End If
            End If
        Next i

    ' Add conditional format
    For i = 2 To lrow
      Cells(i, 10).NumberFormat = "0.00%"
       If Cells(i, 10) > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
        Else
            Cells(i, 11).Interior.ColorIndex = 3
        End If
    Next i

'**********************************************
' 3. Stocks 2015
'**********************************************
'Activate sheet
    Worksheets("2015").Activate
    
        ' Table's titles
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly change"
    Cells(1, 11).Value = "Percent change"
    Cells(1, 12).Value = "Total stock volume"

'Reset values
   volume = 0
   price_final = 0
   price_initial = 0
   change_per = 0
   change = 0
   table_row = 2
   
' Define last row and column
    lrow = Cells(Rows.Count, 2).End(xlUp).Row

       For i = 2 To lrow
' If last observation of the same stock
           If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
' Add volume of each day
               volume = volume + Cells(i, 7).Value
' Extract info if first or last day of the year
                If Cells(i, 2).Value = 20150101 Then
                    price_initial = Cells(i, 3).Value
                    volume_initial = Cells(i, 7).Value
                ElseIf Cells(i, 2).Value = 20151231 Then
                    price_final = Cells(i, 6).Value
                End If
                stock = Cells(i, 1).Value
' Calculate stock values
                change = price_final - price_initial
                If price_initial = 0 Then
                    change_per = 0
                Else
                    change_per = (price_final / price_initial) - 1
                End If
' Print stock values
                Worksheets("2015").Cells(table_row, 9).Value = stock
                Worksheets("2015").Cells(table_row, 10).Value = change
                Worksheets("2015").Cells(table_row, 11).Value = change_per
                Worksheets("2015").Cells(table_row, 12).Value = volume
' Add one to table row
                table_row = table_row + 1
' Reset values
                stock = 0
                change = 0
                change_per = 0
                volume = 0
            Else
' If not last observation of the same stock --> add volume for each day
               volume = volume + Cells(i, 7).Value
' And only extract info if first or last day of the year
                If Cells(i, 2).Value = 20150101 Then
                    price_initial = Cells(i, 3).Value
                ElseIf Cells(i, 2).Value = 20151231 Then
                     price_final = Cells(i, 6).Value
                End If
            End If
        Next i

    ' Add conditional format
    For i = 2 To lrow
      Cells(i, 10).NumberFormat = "0.00%"
       If Cells(i, 10) > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
        Else
            Cells(i, 11).Interior.ColorIndex = 3
        End If
    Next i

'**********************************************
' 4. Stocks 2016
'**********************************************
'Activate sheet
    Worksheets("2015").Activate

    ' Table's titles
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly change"
    Cells(1, 11).Value = "Percent change"
    Cells(1, 12).Value = "Total stock volume"
    
'Reset values
   volume = 0
   price_final = 0
   price_initial = 0
   change_per = 0
   change = 0
   table_row = 2
   
   
' Define last row and column
    lrow = Cells(Rows.Count, 2).End(xlUp).Row

       For i = 2 To lrow
' If last observation of the same stock
           If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
' Add volume of each day
               volume = volume + Cells(i, 7).Value
' Extract info if first or last day of the year
                If Cells(i, 2).Value = 20160101 Then
                    price_initial = Cells(i, 3).Value
                    volume_initial = Cells(i, 7).Value
                ElseIf Cells(i, 2).Value = 20161230 Then
                    price_final = Cells(i, 6).Value
                End If
                stock = Cells(i, 1).Value
' Calculate stock values
                change = price_final - price_initial
                If price_initial = 0 Then
                    change_per = 0
                Else
                    change_per = (price_final / price_initial) - 1
                End If
' Print stock values
                Worksheets("2016").Cells(table_row, 9).Value = stock
                Worksheets("2016").Cells(table_row, 10).Value = change
                Worksheets("2016").Cells(table_row, 11).Value = change_per
                Worksheets("2016").Cells(table_row, 12).Value = volume
' Add one to table row
                table_row = table_row + 1
' Reset values
                stock = 0
                change = 0
                change_per = 0
                volume = 0
            Else
' If not last observation of the same stock --> add volume for each day
               volume = volume + Cells(i, 7).Value
' And only extract info if first or last day of the year
                If Cells(i, 2).Value = 20160101 Then
                    price_initial = Cells(i, 3).Value
                    If price_initial = 0 Then price_initial = 0.0001
                    
                ElseIf Cells(i, 2).Value = 20161230 Then
                     price_final = Cells(i, 6).Value
                End If
            End If
        Next i

    ' Add conditional format
    For i = 2 To lrow
      Cells(i, 10).NumberFormat = "0.00%"
       If Cells(i, 11) > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
        Else
            Cells(i, 11).Interior.ColorIndex = 3
        End If
    Next i

End Sub
