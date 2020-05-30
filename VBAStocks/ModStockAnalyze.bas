Attribute VB_Name = "ModStockAnalyze"
Public Const ConstInc As String = "INC"
Public Const ConstDec As String = "DEC"
Public Const ConstVol As String = "VOL"

''' <summary>
''' This method triggers Stock yearly analysis and greatest analysis
''' and populates those data in all worksheets
''' </summary>
Sub Calculate_Click()

    Dim grtDict As Dictionary
    Dim wsList() As String
    
    ' Get the list of worksheet names
    wsList = GetWorksheetlist
    
    ' Iterate Stock analysis for all worksheets
    For Each Key In wsList
        PopulateYearlyHeaders CStr(Key)
        Set grtDict = PopulateYearlyData(2, 1, CStr(Key))
        PopulateGreatestHeaders CStr(Key)
        PopulateGreatestData grtDict, CStr(Key)
    Next Key
    
    MsgBox ("Calculation Completed!")

End Sub

''' <summary>
''' This method clears yearly and greatest data and formattings
''' </summary>
Sub Clear_Click()
    
    Dim wsList() As String
    
    ' Get the list of worksheets of the workbook
    wsList = GetWorksheetlist
    For Each Key In wsList
        ' Clear the content of early and greatest data
        Worksheets(Key).Range("I:Q").ClearContents
        ' Clear the formatting of early and greatest data
        Worksheets(Key).Range("I:Q").ClearFormats
    Next Key
    MsgBox ("Clear Completed!")
    
End Sub


''' <summary>
''' This method returns the stock details of a specific day
''' when the reference point to the ticker cell is passed
''' </summary>
''' <param name="row"><c>row</c> is the row number</param>
''' <param name="col"><c>col</c> is the new column number</param>
''' <param name="ws"><c>ws</c> is the new worksheet name</param>
Function GetDailyStockDetails(row As Long, col As Integer, ws As String) As CStock
    
    Dim stock As New CStock
    stock.Ticker = Worksheets(ws).Cells(row, col).Value
    stock.StockDate = CLng(Worksheets(ws).Cells(row, col + 1).Value)
    stock.DOpenPrice = CDbl(Worksheets(ws).Cells(row, col + 2).Value)
    stock.DClosePrice = CDbl(Worksheets(ws).Cells(row, col + 5).Value)
    stock.DVolume = CDbl(Worksheets(ws).Cells(row, col + 6).Value)
    Set GetDailyStockDetails = stock
    
End Function

''' <summary>
''' This method populates yearly stock details
''' when the first ticker details reference is passed
''' </summary>
''' <param name="row"><c>row</c> is the row number of the raw data starting</param>
''' <param name="col"><c>col</c> is the column number of the raw data ticker</param>
''' <param name="ws"><c>ws</c> is the worksheet name stock being analysed</param>
Function PopulateYearlyData(row As Long, col As Integer, ws As String) As Dictionary
    
    ' dictionary object to hold Ticker as Key and stock object as value
    Dim stocks As Scripting.Dictionary
    Set stocks = New Scripting.Dictionary
    Dim stock As CStock
    Dim i As Long
    Dim j As Integer
    ' dictionary object to hold greatest stock objects
    Dim greatest As Scripting.Dictionary
    
    ' assign the starting row of stocks
    i = row
    
    ' Looping through all rows which has value for ticker
    Do Until IsEmpty(Worksheets(ws).Cells(i, col).Value)
        Set stock = GetDailyStockDetails(i, col, ws)
        If Not stocks.Exists(stock.Ticker) Then
            ' If the ticker not in the dict, add the object to the dict
            stock.YVolume = stock.DVolume
            stock.YOpenPrice = stock.DOpenPrice
            stock.YClosePrice = stock.DClosePrice
            stocks.Add stock.Ticker, stock
        Else
            ' If the ticker is already in thedict, update the volume and YClosePrice
            stocks.Item(stock.Ticker).YVolume = stocks.Item(stock.Ticker).YVolume + stock.DVolume
            stocks.Item(stock.Ticker).YClosePrice = stock.DClosePrice
        End If
        
        i = i + 1
    Loop
    
    ' Since Yearly data needs to be populated from Raw 2
    j = 2
    
    For Each Key In stocks
        Worksheets(ws).Range("I" & j).Value = stocks(Key).Ticker
        Worksheets(ws).Range("J" & j).Value = stocks(Key).YPriceChange
        
        If (stocks(Key).YPriceChange < 0) Then
            Worksheets(ws).Range("J" & j).Interior.ColorIndex = 3
        ElseIf (stocks(Key).YPriceChange > 0) Then
            Worksheets(ws).Range("J" & j).Interior.ColorIndex = 4
        End If
        
        Worksheets(ws).Range("K" & j).Value = stocks(Key).YPercentChange
        Worksheets(ws).Range("K" & j).NumberFormat = "0.00%"
        Worksheets(ws).Range("L" & j).Value = stocks(Key).YVolume
        j = j + 1
    Next Key
    
    Set greatest = FindGreatest(stocks)
    Set PopulateYearlyData = greatest
    Set stocks = Nothing
    Set greatest = Nothing
End Function

''' <summary>
''' This method finds the Greatest stock details of the year
''' </summary>
Function FindGreatest(yearlyStocks As Dictionary) As Dictionary

    Dim greatest As Scripting.Dictionary
    Set greatest = New Scripting.Dictionary
    Dim grtInc As New CGreatest
    Dim grtDec As New CGreatest
    Dim grtVol As New CGreatest
    
    'Logic to identify  greatest inc, dec, vol stocks
    For Each Key In yearlyStocks
    
        If grtInc.Param < yearlyStocks(Key).YPercentChange Then
            grtInc.Param = yearlyStocks(Key).YPercentChange
            grtInc.Ticker = yearlyStocks(Key).Ticker
        End If
        
        If grtDec.Param > yearlyStocks(Key).YPercentChange Then
            grtDec.Param = yearlyStocks(Key).YPercentChange
            grtDec.Ticker = yearlyStocks(Key).Ticker
        End If
        
        If grtVol.Param < yearlyStocks(Key).YVolume Then
            grtVol.Param = yearlyStocks(Key).YVolume
            grtVol.Ticker = yearlyStocks(Key).Ticker
        End If

    Next Key
    
    greatest.Add ConstInc, grtInc
    greatest.Add ConstDec, grtDec
    greatest.Add ConstVol, grtVol
    
    Set FindGreatest = greatest
    
End Function

''' <summary>
''' This method populates headers
''' when the first header cell reference is passed
''' along with an array of headers
''' </summary>
''' <param name="row"><c>row</c> is the row number header needs to be populated</param>
''' <param name="col"><c>col</c> is the column number 1st header item populated to</param>
''' <param name="ws"><c>ws</c> is the new worksheet name header needs to be populated to</param>
''' <param name="arr"><c>arr</c> is the array contains the column headings</param>
Sub PopulateHeaders(row As Long, col As Integer, ws As String, ByRef arr() As String)
    
    Dim arLength As Integer
    arLength = UBound(arr) - LBound(arr) + 1
    For i = 0 To arLength - 1
        Worksheets(ws).Cells(row, col + i).Value = arr(i)
    Next i

End Sub

''' <summary>
''' This method populates yearly data headers
''' </summary>
Sub PopulateYearlyHeaders(ws As String)

    ' Array to hold yearly headers
    Dim YearlyHeader(3) As String
    YearlyHeader(0) = "Ticker"
    YearlyHeader(1) = "Yearly Change $"
    YearlyHeader(2) = "Percent Change"
    YearlyHeader(3) = "Total Stock Volume"
    
    PopulateHeaders 1, 9, ws, YearlyHeader

End Sub

''' <summary>
''' This method populates Greatest header fields
''' </summary>
Sub PopulateGreatestHeaders(ws As String)

    ' Array to hold greatest headers
    Dim header(2) As String
    header(0) = "Ticker"
    header(1) = "Value"
    
    PopulateHeaders 1, 16, ws, header
    Worksheets(ws).Cells(2, 15).Value = "Greatest % Increase"
    Worksheets(ws).Cells(3, 15).Value = "Greatest % Decrease"
    Worksheets(ws).Cells(4, 15).Value = "Greatest Total Volume"

End Sub

''' <summary>
''' This method populates Greatest fields of the yearly stock prices
''' </summary>
Sub PopulateGreatestData(greatestDic As Dictionary, ws As String)

    Worksheets(ws).Cells(2, 16).Value = greatestDic(ConstInc).Ticker
    Worksheets(ws).Cells(3, 16).Value = greatestDic(ConstDec).Ticker
    Worksheets(ws).Cells(4, 16).Value = greatestDic(ConstVol).Ticker
    Worksheets(ws).Cells(2, 17).Value = greatestDic(ConstInc).Param
    Worksheets(ws).Cells(3, 17).Value = greatestDic(ConstDec).Param
    Worksheets(ws).Cells(4, 17).Value = greatestDic(ConstVol).Param
    
    Worksheets(ws).Cells(2, 17).NumberFormat = "0.00%"
    Worksheets(ws).Cells(3, 17).NumberFormat = "0.00%"

End Sub

''' <summary>
''' This method returns list of worksheets in the current workbook
''' </summary>
Function GetWorksheetlist() As String()

    Dim worksheetList() As String
    ReDim worksheetList(ThisWorkbook.Sheets.Count - 1)

    For i = 1 To ThisWorkbook.Sheets.Count
        worksheetList(i - 1) = ThisWorkbook.Sheets(i).Name
    Next i
    
    GetWorksheetlist = worksheetList

End Function


