Attribute VB_Name = "WesabeAutomation"
Option Explicit

Public Sub ClearWesabeTransactionData()
    Dim transactionsBook As Workbook
    Dim transactionsSheet As Worksheet
    
    Set transactionsBook = ThisWorkbook
    Set transactionsSheet = transactionsBook.Sheets(Evaluate(ThisWorkbook.Names("TransactionsSheetName").Value))
    
    Dim transactionsCleared As Boolean
    If Not transactionsSheet Is Nothing Then
        
        Dim transactionsQuery As QueryTable
        On Error Resume Next
        Set transactionsQuery = transactionsSheet.QueryTables("TransactionsQuery")
        
        If Not transactionsQuery Is Nothing Then
            transactionsQuery.ResultRange.ClearContents
        End If
    End If

    If transactionsSheet Is Nothing Then
        MsgBox "Could not locate the transactions worksheet.", vbExclamation, "ClearWesabeTransactionsData Error"
    ElseIf transactionsQuery Is Nothing Then
        MsgBox "Could not locate the transactions web query.", vbExclamation, "ClearWesabeTransactionsData Error"
    End If
End Sub

Public Sub ClearWesabeTransactionsXmlRangeNames()
    Dim transactionsBook As Workbook
    Dim transactionsSheet As Worksheet
    
    Set transactionsBook = ThisWorkbook
    Set transactionsSheet = transactionsBook.Sheets(Evaluate(ThisWorkbook.Names("TransactionsSheetName").Value))
    
    Dim rangeNamePrefix As String
    rangeNamePrefix = transactionsSheet.Name & "!"
    If Not transactionsBook Is Nothing Then
        Dim calculationSetting As Variant
        calculationSetting = Application.Calculation
        Application.Calculation = xlCalculationManual
        
        On Error Resume Next 'Continue on error so that the calculation setting is always restored
        
        Dim namedRange As Name
        Dim i As Integer
        For i = transactionsBook.Names.Count To 1 Step -1
            Set namedRange = transactionsBook.Names.Item(i)
            If Left(namedRange.Name, Len(rangeNamePrefix & "WesabeTransactionDate")) = rangeNamePrefix & "WesabeTransactionDate" _
            Or Left(namedRange.Name, Len(rangeNamePrefix & "WesabeTransactionWeek")) = rangeNamePrefix & "WesabeTransactionWeek" _
            Or Left(namedRange.Name, Len(rangeNamePrefix & "WesabeAmount")) = rangeNamePrefix & "WesabeAmount" _
            Or Left(namedRange.Name, Len(rangeNamePrefix & "WesabeAggregateAmount")) = rangeNamePrefix & "WesabeAggregateAmount" _
            Or Left(namedRange.Name, Len(rangeNamePrefix & "WesabeTagName")) = rangeNamePrefix & "WesabeTagName" _
            Or Left(namedRange.Name, Len(rangeNamePrefix & "WesabeSplitAmount")) = rangeNamePrefix & "WesabeSplitAmount" Then
                namedRange.Delete
            End If
        Next i
        
        On Error GoTo 0
        
        Application.Calculate
        Application.Calculation = calculationSetting
    End If
End Sub

Public Sub DownloadTransactionsFromWesabe()
Attribute DownloadTransactionsFromWesabe.VB_Description = "Downloads the latest transaction data from Wesabe and updates the spreadsheet calculations."
Attribute DownloadTransactionsFromWesabe.VB_ProcData.VB_Invoke_Func = "D\n14"
    Dim transactionsBook As Workbook
    Dim transactionsSheet As Worksheet
    
    Set transactionsBook = ThisWorkbook
    Set transactionsSheet = transactionsBook.Sheets(Evaluate(ThisWorkbook.Names("TransactionsSheetName").Value))
    
    Dim transactionsRefreshed As Boolean
    If Not transactionsSheet Is Nothing Then
        
        Dim transactionsQuery As QueryTable
        On Error Resume Next
        Set transactionsQuery = transactionsSheet.QueryTables("TransactionsQuery")
        
        If Not transactionsQuery Is Nothing Then
            transactionsRefreshed = transactionsQuery.Refresh(False)
            If transactionsRefreshed Then
                RedefineWesabeTransactionsXmlRangeNames
            End If
        End If
    End If

    If transactionsSheet Is Nothing Then
        MsgBox "Could not locate the transactions worksheet.", vbExclamation, "DownloadTransactionsFromWesabe Error"
    ElseIf transactionsQuery Is Nothing Then
        MsgBox "Could not locate the transactions web query.", vbExclamation, "DownloadTransactionsFromWesabe Error"
    ElseIf Not transactionsRefreshed Then
        MsgBox "Could not refresh the transactions data.", vbExclamation, "DownloadTransactionsFromWesabe Error"
    End If
End Sub

Public Sub RedefineWesabeTransactionsXmlRangeNames()
Attribute RedefineWesabeTransactionsXmlRangeNames.VB_Description = "Organizes downloaded transactions by date to allow for efficient calculations."
Attribute RedefineWesabeTransactionsXmlRangeNames.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim transactionsBook As Workbook
    Dim transactionsSheet As Worksheet
    
    Set transactionsBook = ThisWorkbook
    Set transactionsSheet = transactionsBook.Sheets(Evaluate(ThisWorkbook.Names("TransactionsSheetName").Value))
    
    If Not transactionsSheet Is Nothing Then
        Dim calculationSetting As Variant
        calculationSetting = Application.Calculation
        Application.Calculation = xlCalculationManual
        
        On Error Resume Next 'Continue on error so that the calculation setting is always restored
        DefineWesabeTransactionsRangeNamesXml transactionsSheet
        On Error GoTo 0
        
        Application.Calculate
        Application.Calculation = calculationSetting
    End If
End Sub

Private Sub DefineWesabeTransactionsRangeNamesXml(transactionsSheet As Worksheet)
    With transactionsSheet
        Dim transactionDates As Range
        Dim transactionDatesColumn As Integer
        Dim amountsColumn As Integer
        Dim aggregateAmountsColumn As Integer
        Dim tagsColumn As Integer
        Dim splitAmountsColumn As Integer
        
        Dim headerRow As Range
        'Skip row 1 for root "/txactions" element
        Set headerRow = .Range(.UsedRange.Cells(2, 1).Address, .UsedRange.Cells(2, .UsedRange.Columns.Count).Address)
        Dim headerCell As Range
        For Each headerCell In headerRow
            Select Case headerCell.Value
                Case "/txaction/date":
                    Set transactionDates = .Range(headerCell.Offset(1, 0), .Cells(.Rows.Count, headerCell.Column).End(xlUp))
                    transactionDatesColumn = headerCell.Column
                
                Case "/txaction/amount":
                    amountsColumn = headerCell.Column
                
                Case "/txaction/tags/tag/split-amount":
                    splitAmountsColumn = headerCell.Column
                
                Case "/txaction/tags/tag/name":
                    tagsColumn = headerCell.Column
                
                Case "/txaction/amount/#agg":
                    aggregateAmountsColumn = headerCell.Column
            End Select
        Next
        
        PartitionWesabeData _
            transactionsSheet, _
            transactionDatesColumn, _
            amountsColumn, _
            aggregateAmountsColumn, _
            tagsColumn, _
            splitAmountsColumn, _
            transactionDates
    End With
End Sub
        
Private Sub PartitionWesabeData(transactionsSheet As Worksheet, transactionDatesColumn As Integer, amountsColumn As Integer, aggregateAmountsColumn As Integer, tagsColumn As Integer, splitAmountsColumn As Integer, transactionDates As Range)
        Dim minTransDate As Date
        Dim maxTransDate As Date
        minTransDate = Application.WorksheetFunction.Min(transactionDates)
        maxTransDate = Application.WorksheetFunction.Max(transactionDates)
        
        Dim transMonthCount As Integer
        transMonthCount = MonthsBetweenDates(minTransDate, maxTransDate)
        
        Dim firstCellInMonth() As Integer
        Dim lastCellInMonth() As Integer
        ReDim firstCellInMonth(0 To transMonthCount)
        ReDim lastCellInMonth(0 To transMonthCount)

        Dim transIsoWeekCount As Integer
        transIsoWeekCount = IsoWeeksBetweenDates(minTransDate, maxTransDate)
        
        Dim firstCellInIsoWeek() As Integer
        Dim lastCellInIsoWeek() As Integer
        ReDim firstCellInIsoWeek(0 To transIsoWeekCount)
        ReDim lastCellInIsoWeek(0 To transIsoWeekCount)

        Dim dateCell As Range
        Dim cellMonthIndex As Integer
        Dim cellIsoWeekIndex As Integer
        For Each dateCell In transactionDates.Cells
            cellMonthIndex = MonthsBetweenDates(minTransDate, CDate(dateCell.Value))
            If firstCellInMonth(cellMonthIndex) = 0 Then
                firstCellInMonth(cellMonthIndex) = dateCell.Row
            End If
            lastCellInMonth(cellMonthIndex) = dateCell.Row
        
            cellIsoWeekIndex = IsoWeeksBetweenDates(minTransDate, CDate(dateCell.Value))
            If firstCellInIsoWeek(cellIsoWeekIndex) = 0 Then
                firstCellInIsoWeek(cellIsoWeekIndex) = dateCell.Row
            End If
            lastCellInIsoWeek(cellIsoWeekIndex) = dateCell.Row
        Next dateCell
        
        CreateWesabeRangeNames _
            transactionsSheet, _
            transactionDatesColumn, _
            amountsColumn, _
            aggregateAmountsColumn, _
            tagsColumn, _
            splitAmountsColumn, _
            transMonthCount, _
            transIsoWeekCount, _
            firstCellInMonth, _
            lastCellInMonth, _
            firstCellInIsoWeek, _
            lastCellInIsoWeek
End Sub
        
Private Sub CreateWesabeRangeNames(transactionsSheet As Worksheet, transactionDatesColumn As Integer, amountsColumn As Integer, aggregateAmountsColumn As Integer, tagsColumn As Integer, splitAmountsColumn As Integer, transMonthCount As Integer, transIsoWeekCount As Integer, firstCellInMonth() As Integer, lastCellInMonth() As Integer, firstCellInIsoWeek() As Integer, lastCellInIsoWeek() As Integer)
    With transactionsSheet
        Dim firstDateCell As Range
        Dim monthRangeIndex As Integer
        For monthRangeIndex = 0 To transMonthCount
            Set firstDateCell = .Cells(firstCellInMonth(monthRangeIndex), transactionDatesColumn)
            Dim monthNameSuffix As String
            monthNameSuffix = Year(firstDateCell.Value) & Month(firstDateCell.Value)
            
            If transactionDatesColumn > 0 Then
                .Names.Add _
                    Name:="WesabeTransactionDate" & monthNameSuffix, _
                    RefersTo:="=" & .Cells(firstCellInMonth(monthRangeIndex), transactionDatesColumn).Address & ":" & .Cells(lastCellInMonth(monthRangeIndex), transactionDatesColumn).Address
            End If
            
            If amountsColumn > 0 Then
                .Names.Add _
                    Name:="WesabeAmount" & monthNameSuffix, _
                    RefersTo:="=" & .Cells(firstCellInMonth(monthRangeIndex), amountsColumn).Address & ":" & .Cells(lastCellInMonth(monthRangeIndex), amountsColumn).Address
            End If
            
            If aggregateAmountsColumn > 0 Then
                .Names.Add _
                    Name:="WesabeAggregateAmount" & monthNameSuffix, _
                    RefersTo:="=" & .Cells(firstCellInMonth(monthRangeIndex), aggregateAmountsColumn).Address & ":" & .Cells(lastCellInMonth(monthRangeIndex), aggregateAmountsColumn).Address
            End If
            
            If tagsColumn > 0 Then
                .Names.Add _
                    Name:="WesabeTagName" & monthNameSuffix, _
                    RefersTo:="=" & .Cells(firstCellInMonth(monthRangeIndex), tagsColumn).Address & ":" & .Cells(lastCellInMonth(monthRangeIndex), tagsColumn).Address
            End If
            
            If splitAmountsColumn > 0 Then
                .Names.Add _
                    Name:="WesabeSplitAmount" & monthNameSuffix, _
                    RefersTo:="=" & .Cells(firstCellInMonth(monthRangeIndex), splitAmountsColumn).Address & ":" & .Cells(lastCellInMonth(monthRangeIndex), splitAmountsColumn).Address
            End If
        Next monthRangeIndex
    
        Dim isoWeekRangeIndex As Integer
        For isoWeekRangeIndex = 0 To transIsoWeekCount
            Set firstDateCell = .Cells(firstCellInIsoWeek(isoWeekRangeIndex), transactionDatesColumn)
            Dim weekNameSuffix As String
            weekNameSuffix = IsoYear(firstDateCell.Value) & "W" & IsoWeek(firstDateCell.Value)
            
            If transactionDatesColumn > 0 Then
                .Names.Add _
                    Name:="WesabeTransactionDate" & weekNameSuffix, _
                    RefersTo:="=" & .Cells(firstCellInIsoWeek(isoWeekRangeIndex), transactionDatesColumn).Address & ":" & .Cells(lastCellInIsoWeek(isoWeekRangeIndex), transactionDatesColumn).Address
            End If
        
            If amountsColumn > 0 Then
                .Names.Add _
                    Name:="WesabeAmount" & weekNameSuffix, _
                    RefersTo:="=" & .Cells(firstCellInIsoWeek(isoWeekRangeIndex), amountsColumn).Address & ":" & .Cells(lastCellInIsoWeek(isoWeekRangeIndex), amountsColumn).Address
            End If
            
            If aggregateAmountsColumn > 0 Then
                .Names.Add _
                    Name:="WesabeAggregateAmount" & weekNameSuffix, _
                    RefersTo:="=" & .Cells(firstCellInIsoWeek(isoWeekRangeIndex), aggregateAmountsColumn).Address & ":" & .Cells(lastCellInIsoWeek(isoWeekRangeIndex), aggregateAmountsColumn).Address
            End If
            
            If tagsColumn > 0 Then
                .Names.Add _
                    Name:="WesabeTagName" & weekNameSuffix, _
                    RefersTo:="=" & .Cells(firstCellInIsoWeek(isoWeekRangeIndex), tagsColumn).Address & ":" & .Cells(lastCellInIsoWeek(isoWeekRangeIndex), tagsColumn).Address
            End If
            
            If splitAmountsColumn > 0 Then
                .Names.Add _
                    Name:="WesabeSplitAmount" & weekNameSuffix, _
                    RefersTo:="=" & .Cells(firstCellInIsoWeek(isoWeekRangeIndex), splitAmountsColumn).Address & ":" & .Cells(lastCellInIsoWeek(isoWeekRangeIndex), splitAmountsColumn).Address
            End If
        Next
    End With
End Sub

Private Function MonthsBetweenDates(startDate As Date, endDate As Date) As Integer
    MonthsBetweenDates = ((Year(endDate) - Year(startDate)) * 12) + (Month(endDate) - Month(startDate))
End Function

Private Function IsoWeeksBetweenDates(startDate As Date, endDate As Date) As Integer
    Dim dayDiff As Integer
    dayDiff = DateValue(endDate) - DateValue(startDate)
    Dim weekdayDiff As Integer
    weekdayDiff = Weekday(endDate, vbMonday) - Weekday(startDate, vbMonday)
    IsoWeeksBetweenDates = ((dayDiff + ((dayDiff Mod 7) - weekdayDiff)) \ 7)
End Function

Public Function IsoWeek(d1 As Date) As Integer
    ' Copied from http://msdn2.microsoft.com/en-us/library/bb277364.aspx
    ' Provided by Daniel Maher.
    Dim d2 As Long
    d2 = DateSerial(Year(d1 - Weekday(d1 - 1) + 4), 1, 3)
    IsoWeek = Int((d1 - d2 + Weekday(d2) + 5) / 7)
End Function

Public Function IsoYear(d1 As Date) As Integer
    'Compute the closest Thursday
    Dim d2 As Date
    d2 = d1 + (4 - Weekday(d1, vbMonday))
    IsoYear = Year(d2)
End Function
