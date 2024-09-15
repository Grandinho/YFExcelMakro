Attribute VB_Name = "Modul3"
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
        url = "https://query1.finance.yahoo.com/v8/finance/chart/" & Symbol & "?period1=" & period1 & "&period2=" & period2 & "&interval=1d&events=split"
        
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
            
            If result.Exists("events") Then
                If result("events").Exists("splits") Then
                    Dim splits As Object
                    Set splits = result("events")("splits")
                    
                    Dim key As Variant
                    For Each key In splits.Keys
                        Dim unixTimestamp As Long
                        Dim splitDate As Date
                        Dim splitYear As Integer
                        
                        unixTimestamp = CLng(key)
                        splitDate = DateAdd("s", unixTimestamp, #1/1/1970#)
                        splitYear = year(splitDate)
                        Dim dateCell As Range
                        Dim findString As String
                        findString = "01.01." & splitYear
                        Set dateCell = ws.Range(startCellDates + ":" + endCellDates).Find(What:=findString)
                        If Not dateCell Is Nothing Then
                            ws.Cells(dateCell.row, cell.Column).Value = "True"
                        End If
                        
                    Next key
                End If
            End If
            
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



