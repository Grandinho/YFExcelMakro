Attribute VB_Name = "Modul1"
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
        symbol = cell.Value
        
        Debug.Print cell.Column
        Debug.Print Col_Letter(cell.Row)
        
        ' Construct the URL
        url = "https://query1.finance.yahoo.com/v8/finance/chart/" & symbol & "?period1=" & period1 & "&period2=" & period2 & "&interval=1mo&events=history"
        
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
    
            
            
            Dim i As Long
            Dim cellResult As String
            If result("indicators")("quote")(1).Exists("close") Then
                If result("indicators")("quote")(1)("close").count > 0 Then
                    For i = 1 To result("indicators")("quote")(1)("close").count
                        Dim unixTimestamp As Long
                        Dim entryDate As Date
                        cellResult = result("indicators")("quote")(1)("close")(i)
                        unixTimestamp = result("timestamp")(i)
                        entryDate = DateAdd("d", 1, UnixToDate(unixTimestamp))
                       ' With ws.Range(startCellDates + ":" + endCellDates)
                       '     Dim rowCell As Range
                       '     Set rowCell = .Find(DateSerial(Year(entryDate), Month(entryDate), 1), LookIn:=xlValues)
                       '     If Not rowCell Is Nothing Then
                       '         ws.Range(Col_Letter(cell.Column) + CStr(rowCell.Row)).NumberFormat = "@"
                       '         ws.Range(Col_Letter(cell.Column) + CStr(rowCell.Row)).Value = cellResult
                       '     End If
                       ' End With
                        For Each rowCell In ws.Range(startCellDates + ":" + endCellDates).Cells
                            ws.Range(Col_Letter(cell.Column) + CStr(rowCell.Row)).NumberFormat = "@"
                            Dim cellDate As Date
                            cellDate = CDate(rowCell.Value)
                            If Year(cellDate) = Year(entryDate) And Month(cellDate) = Month(entryDate) Then
                                 ws.Range(Col_Letter(cell.Column) + CStr(rowCell.Row)).Value = cellResult
                                 GoTo NextIteration
                            End If
                         Next
NextIteration:
                    Next i
                End If
            End If
        Else
            ' MsgBox "Failed for " + symbol + " to retrieve data. Status code: " & httpRequest.Status, vbExclamation
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

