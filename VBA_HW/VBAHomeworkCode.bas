Attribute VB_Name = "Module1"
Sub Stock()

    'Varaiables Ticker, Total Stock Volume, SummaryTableRow, LastRow, r NumSheet
    Dim Ticker As String
    'Make sure TSV is long enough - huge number
    Dim TotalStockVolume As Long
    Dim SummaryTableRow As String
    'LastRow is big number - make long
   ' Make sure r is long too
    Dim r As Long
    Dim NumSheet As Worksheet
    
    Dim LastRow As Long
    'Loop time
    'need loop for each different sheet in workbook
    For Each NumSheet In Worksheets
        NumSheet.Activate
        LastRow = NumSheet.Range("A1", NumSheet.Range("A1").End(xlDown)).Rows.Count
    'Add the headers and places
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Stock Volume"

    'Assign Values
        TotalStockVolume = 0
        SummaryTableRow = 2

    'Start the Loop by using last row
      ' LastRow = Cells(Rows.Count, "A").End(xIUp).Row

    'Make Loop
        For r = 2 To LastRow

    'Begin if statement. Utilizes conditionals for comparing cells.
            If Cells(r + 1).Value <> Cells(r, 1).Value Then

    'Define TotalStockValue
                TotalStockVolume = TotalStockVolume + Cells(r, 7).Value

    'Show Ticker values in summary
                Range("I" & SummaryTableRow).Value = Ticker

    'Show the TotalStock Volume values in summary
                Range("J" & SummaryTableRow).Value = TotalStockValue

    'Add one to Summary and set Total Stock Volume back to 0
                SummaryTableRow = SummaryTableRow + 1
                TotalStockVolume = 0

    'But if above statement is not true...
            Else
                TotalStockVolume = TotalStockVolume + Cells(r, 7).Value

            End If
        Next r
    Next NumSheet


End Sub


