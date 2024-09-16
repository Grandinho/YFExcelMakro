Attribute VB_Name = "Modul2"
Sub GetYahooFinanceData()
    Dim ws As Worksheet
    Dim startCellSymbol As String, endCellSymbol As String
    Dim startCellDates As String, endCellDates As String
    Dim url As String
    Dim httpRequest As Object
    Dim jsonText As String
    Dim jsonObject As Object
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    Dim count As Integer
    
    count = Range("A3", "A14").Cells.count
    
    Dim count2 As Integer
    ' Dim Range1 As Range
    ' Dim Range2 As Range
    ' Range1 = Range("B2")
    ' Range2 = Range("F2")
    ' count2 = Range("B2", "F2").Cells.count
    

    
    ' Get user inputs from cells
    startCellSymbol = InputBox("Input cell where first symbol is placed")
    endCellSymbol = InputBox("Input cell where last symbol is placed")
    
    startCellDates = InputBox("Input cell with start date")
    endCellDates = InputBox("Input cell with end date")
    
    
    period1 = DateToUnixTimestamp(ws.Range(startCellDates).Value)
    period2 = DateToUnixTimestamp(ws.Range(endCellDates).Value)
    
    For Each cell In ws.Range(startCellSymbol + ":" + endCellSymbol).Cells
        Symbol = cell.Value
        
        Debug.Print cell.Column
        Debug.Print Col_Letter(cell.row)
        
        ' Construct the URL
        url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Symbol & "?period1=" & period1 & "&period2=" & period2 & "&interval=1d&events=div"
        
        ' Create HTTP request object
        Set httpRequest = CreateObject("MSXML2.XMLHTTP")
        
        ' Send the request
        With httpRequest
            .Open "GET", url, False
            .Send
        End With
        
        ' Check if the request was successful
        If httpRequest.Status = 200 Then
            jsonText = httpRequest.responseText
            
            ' Parse JSON
            Set jsonObject = JsonConverter.ParseJson(jsonText)
            
            ' Extract and populate data
            Dim result As Object
            Set result = jsonObject("chart")("result")(1)
            
            ' Create a dictionary to store dividends by year
            Dim dividendsByMonth As Object
            Set dividendsByMonth = CreateObject("Scripting.Dictionary")
            
            ' Process dividends
            ' And result("events").Exists("dividends")
            If result.Exists("events") Then
                If result("events").Exists("dividends") Then
                    Dim dividends As Object
                    Set dividends = result("events")("dividends")
                    
                    Dim key As Variant
                    For Each key In dividends.Keys
                        Dim unixTimestamp As Long
                        Dim dividendDate As Date
                        Dim dividendAmount As Double
                        Dim dividendYear As Integer
                        Dim dividendMonth As Integer
                        Dim keyString As String
                        
                        unixTimestamp = CLng(key)
                        dividendDate = DateAdd("s", unixTimestamp, #1/1/1970#)
                        dividendAmount = dividends(key)("amount")
                        dividendYear = Year(dividendDate)
                        dividendMonth = Month(dividendDate)
                        
                        keyString = "01." + "0" + CStr(dividendMonth) + "." + CStr(dividendYear)
                        If dividendsByMonth.Exists(keyString) Then
                            dividendsByMonth(keyString) = dividendsByMonth(keyString) + dividendAmount
                        Else
                            dividendsByMonth.Add keyString, dividendAmount
                        End If
                    Next key
                End If
            End If
            
            ' Write dividends to Excel
            Dim year2 As Variant
            
            For Each year2 In dividendsByMonth.Keys
                Dim dateCell As Range
                Dim findDate As Date
                findDate = CDate(year2)
                Set dateCell = ws.Range(startCellDates + ":" + endCellDates).Find(What:=findDate)
                If Not dateCell Is Nothing Then
                    ws.Cells(dateCell.row, cell.Column).Value = dividendsByMonth(year2)
                End If
            Next year2
            
        Else
            ' MsgBox "Failed to retrieve data. Status code: " & httpRequest.Status, vbExclamation
        End If
        
        ' Clean up
        Set httpRequest = Nothing
        Set jsonObject = Nothing
    Next
End Sub

Function DateToUnixTimestamp(dateValue As Date) As Long
    DateToUnixTimestamp = CLng((dateValue - #1/1/1970#) * 86400)
End Function

Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Function UnixToDate(unixTimestamp As Long) As Date
    UnixToDate = DateAdd("s", unixTimestamp, #1/1/1970#)
End Function


