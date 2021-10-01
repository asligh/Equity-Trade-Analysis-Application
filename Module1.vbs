Attribute VB_Name = "Module1"
'Author: Ashley Sligh

'Date: 9/20/2021

'Purpose: Allows the decision maker to quickly access "price change", "yearly % change",
'         and "yearly trade volume" for given publicly traded equity securities over the
'         course of a 262 or 261 business day period. The program also provides the summary
'         statistics, "greatest % increase", "greatest % decrease", and "total trade volume"
'         for the same period.  If the timeperiod for a security is less than 261 days, data
'         analysis will still be performed.

'Assumptions: 1. The program assumes that, "ticker" is a required field and will always be populated
'             2. If the opening price on the first business day of the year is 0, then the opening price
'                on the 2nd business day will contain a non-zero value
'             3. The worksheet does not contain hidden rows
'             4. End-user has at least quad-core processing capability or comparible for computational efficiency
'

'Limitations: 1. This solution does not scale and is only for relatively small datasets. As records grow into
'                the millions, a different approach would be needed.
'             2. Conditional formatting is not applied in the formal sense in order to optimize performance.
'                Cell shading is done as a workaround.

'Dependencies: Data is static, no dependencies

'Stored Procedures: None

'Referenced Tables: None

'------------------------------------------'
'                                          '
'                 OPTIONS                  '
'                                          '
'------------------------------------------'

Option Explicit
Option Base 1

'------------------------------------------'
'                                          '
'                CONSTANTS                 '
'                                          '
'------------------------------------------'

Public Const NUMBER_WORKING_DAYS As Integer = 262     'Stores the number of non-leap year business days in a Gregorian Calendar year
Public Const NUMBER_OF_RAW_DATA_COLS As Integer = 7   'Stores the number of columns to process
Public Const ANALYSIS_COL_COUNT As Integer = 4        'Stores the number of summary analysis columns
Public Const SUMMARY_COL_COUNT As Integer = 3         'Stores the number of summary statistics columns
Public Const SUMMARY_HEADER_COUNT As Integer = 4      'Stores the number of summary statistics row headings

Public Const TICKER_IDX As Integer = 1          'Stores the column location of ticker
Public Const TRADE_DATE_IDX As Integer = 2      'Stores the column location of trade date
Public Const OPEN_QUOTE_IDX As Integer = 3      'Stores the column location of open quote
Public Const HIGH_QUOTE_IDX As Integer = 4      'Stores the column location of high quote
Public Const LOW_QUOTE_IDX As Integer = 5       'Stores the column location of low quote
Public Const CLOSE_QUOTE_IDX As Integer = 6     'Stores the column location of close quote
Public Const VOLUME_IDX As Integer = 7          'Stores the column location of volume

'------------------------------------------'
'                                          '
'                  GLOBAL                  '
'                 VARIABLES                '
'                                          '
'------------------------------------------'

Global analysis_columns(1 To ANALYSIS_COL_COUNT) As String       'Stores the number of raw data columns to process
Global summary_columns(1 To SUMMARY_COL_COUNT) As String         'Stores the number of analysis columns to process
Global summary_row_headers(1 To SUMMARY_HEADER_COUNT) As String  'Stores the number of statistical analysis header rows to process

Global greatest_pct_increase(1 To 2) As Variant  'Stores the greatest percentage increase summary statistic
Global greatest_pct_decrease(1 To 2) As Variant  'Stores the greatest percentage decrease summary statistic
Global greatest_tot_vol(1 To 2) As Variant       'Stores the greatest total volume summary statistic

Global home As String   'Stores handle to begining of worksheet

Global analysis_col_target As String           'Stores start of summary analysis table
Global analysis_summary_col_target As String   'Stores summary column cell
Global analysis_summary_row_target As String   'Stores summary row cell

Global analysis_row_idx As Integer   'Stores index analysis row
Global analysis_col_idx  As Integer  'Stores index of analysis column

Global security_data(1 To NUMBER_WORKING_DAYS, 1 To NUMBER_OF_RAW_DATA_COLS) As String  'Stores 1-year trade record batch

Global total_trade_volume As Variant  'For keeping running total of trade volume on a per-security basis

Global trades_placeholder As String   'Stores range address of new trade batch

'------------------------------------------'
'                                          '
'               SUB-ROUTINES               '
'                                          '
'------------------------------------------'

'Description: Start of program
'Parameters: None
'Returns: N/A
Sub Run_Program()

    Dim sheet_idx As Integer
    Dim worksheet_record_count As Integer
    
    worksheet_record_count = ThisWorkbook.Sheets.count
    
    InitializeGlobalVars
    
    For sheet_idx = 1 To worksheet_record_count
        Sheets(sheet_idx).Select
        InitializeWksheetVars (sheet_idx)
        Process_Security_Batches (worksheet_record_count)
        WriteAnalysisSummary
        Range(home).Select
    Next sheet_idx
    
    Sheets(1).Select
    MsgBox ("All Done!")
    
End Sub

'------------------------------------------'
'                                          '
'                FUNCTIONS                 '
'                                          '
'------------------------------------------'

'Description: Initialize global variables before data processing begins
'Parameters: None
'Returns: Function does not return a value
Private Function InitializeGlobalVars()

    'Running total of security trade volume
    total_trade_volume = 0
    
    'Trade process placeholder
    trades_placeholder = ""
    
    'Initialize range targets
    home = "A2"
    analysis_col_target = "I1"
    analysis_summary_col_target = "O1"
    analysis_summary_row_target = "N2"
    
    'Set analysis column headers
    analysis_columns(1) = "Ticker"
    analysis_columns(2) = "Yearly Change"
    analysis_columns(3) = "Percent Change"
    analysis_columns(4) = "Total Stock Volume"
    
    'Set summary analysis column headers
    summary_columns(1) = "Ticker"
    summary_columns(2) = "Value"
    
    'Set summary analysis row headers
    summary_row_headers(1) = "Greatest % Increase"
    summary_row_headers(2) = "Greatest % Decrease"
    summary_row_headers(3) = "Greatest Total Volume"
    
End Function

'Description: Initialize worksheet variables
'Parameters: None
'Returns: Function does not return a value
Private Function InitializeWksheetVars(ByVal sheet_idx As Integer)

    Erase security_data
    Erase greatest_pct_increase
    Erase greatest_pct_decrease
    Erase greatest_tot_vol
    
    analysis_row_idx = 2
    analysis_col_idx = 9
    total_trade_volume = 0
    trades_placeholder = ""
    Sheets(sheet_idx).Select
    WriteHeaderData analysis_columns, analysis_col_target, True
    WriteHeaderData summary_columns, analysis_summary_col_target, True
    WriteHeaderData summary_row_headers, analysis_summary_row_target, False
    ActiveSheet.Range(home).Select
    
End Function

'Description: Write all worksheet header data
'Parameters: 1. ByRef columns() As String       - Data to be written
'            2. ByVal target As String          - Target location of write
'            3. ByVal isColumnHeader As Boolean - Direction to write data (either as column header or row header) as expressed by True or False
'Returns: Function does not return a value
Private Function WriteHeaderData(ByRef columns() As String, ByVal target As String, ByVal isColumnHeader As Boolean)

    Dim col_count As Integer
    Dim idx As Integer

    'https://www.get-digital-help.com/how-to-use-the-lbound-and-ubound-functions/
    col_count = UBound(columns) - LBound(columns)

    Range(target).Select 'Prep the write
    
    For idx = 1 To col_count
    
        With ActiveCell
        
            .Value = columns(idx)
            .ColumnWidth = Len(ActiveCell.Value)
            
            If isColumnHeader Then
                .Offset(0, 1).Activate
            Else
               .Offset(1, 0).Activate
            End If
        
        End With
        
    Next idx

End Function

'Description: Sargeant method controlling record processing across current worksheet and all other worksheets
'Parameters: ByVal worksheet_record_count As Integer - The total number of worksheets to be processed
'Returns: Function does not return a value
Private Function Process_Security_Batches(ByVal worksheet_record_count As Integer)

    Dim security_record_count As Integer
    Dim security_col_count As Integer
    Dim current_row_num As Integer
    Dim current_ticker As String
    
    Dim array_row_idx As Integer
    
    array_row_idx = 1
    
    current_ticker = ActiveCell.Value 'Priming read
    
    'Much debate around how to effectively obtain a row count:
    'https://stackoverflow.com/questions/11169445/error-in-finding-last-used-cell-in-excel-with-vba
    
    'Processing regimine avoids hidden row pitfalls associated with techniques to obtain row count
    While current_ticker <> ""
         
        While current_ticker = ActiveCell.Value
        
            Read_Security_Record security_data, array_row_idx
            array_row_idx = array_row_idx + 1
        
        Wend
        
        trades_placeholder = ActiveCell.Address
        AnalyzeRecords security_data
        Erase security_data
        total_trade_volume = 0
        array_row_idx = 1
        current_ticker = ActiveCell.Value
        
    Wend

End Function

'Description: Reads and persists current equity security record into volitile memory
'Parameters: ByRef security_data() As String    - One batch of security records
'            ByRef array_row_idx As Integer     - Location of security data to be stored in volatile memory
'Returns: Function does not return a value
Private Function Read_Security_Record(ByRef security_data() As String, ByRef array_row_idx As Integer)

    'Declaration
    Dim ticker_val As String
    Dim trade_date_val  As String
    Dim open_quote_val  As String
    Dim high_quote_val  As String
    Dim low_quote_val  As String
    Dim close_quote_val  As String
    Dim vol_val As String
    Dim running_vol_count As String

    'Initialization
    ticker_val = ""
    trade_date_val = ""
    open_quote_val = ""
    high_quote_val = ""
    low_quote_val = ""
    close_quote_val = ""
    
    'Local variable value assignment less 1 for zero-based cell referencing in this case
    With ActiveCell
    
        ticker_val = .Value                                        '1
        trade_date_val = .Offset(0, TRADE_DATE_IDX - 1).Value      '2
        open_quote_val = .Offset(0, OPEN_QUOTE_IDX - 1).Value      '3
        high_quote_val = .Offset(0, HIGH_QUOTE_IDX - 1).Value      '4
        low_quote_val = .Offset(0, LOW_QUOTE_IDX - 1).Value        '5
        close_quote_val = .Offset(0, CLOSE_QUOTE_IDX - 1).Value    '6
        vol_val = .Offset(0, VOLUME_IDX - 1).Value                 '7
    
    End With
    
    'Persistance
    security_data(array_row_idx, TICKER_IDX) = ticker_val
    security_data(array_row_idx, TRADE_DATE_IDX) = trade_date_val
    security_data(array_row_idx, OPEN_QUOTE_IDX) = open_quote_val
    security_data(array_row_idx, HIGH_QUOTE_IDX) = high_quote_val
    security_data(array_row_idx, LOW_QUOTE_IDX) = low_quote_val
    security_data(array_row_idx, CLOSE_QUOTE_IDX) = close_quote_val
    security_data(array_row_idx, VOLUME_IDX) = vol_val
    
    total_trade_volume = CLng(vol_val) + total_trade_volume 'Running count of daily trade volume
    
    ActiveCell.Offset(1, 0).Select
    
End Function

'Description: Performs analysis on the current batch of equity security records and writes analysis to current worksheet
'Parameters: ByRef security_data() As String  - 1 batch of equity security records
'Returns: Function does not return a value
Private Function AnalyzeRecords(ByRef security_data() As String)

    Dim ticker_val As String
    Dim record_count As Integer
    Dim open_quote_val As Double
    Dim last_business_day_close_quote As Double
    Dim yearly_change As Double
    Dim percent_change As Double

    record_count = GetRecordCount(security_data) 'Record count is the number's worth of business days quotes were processed for
    
    '(1) Analyze Data---------------------------------------------------------------
    
    If record_count = NUMBER_WORKING_DAYS Then
    
        last_business_day_close_quote = CDbl(security_data(NUMBER_WORKING_DAYS, CLOSE_QUOTE_IDX))
        
    ElseIf record_count = (NUMBER_WORKING_DAYS - 1) Then 'Leap year
    
        last_business_day_close_quote = CDbl(security_data(NUMBER_WORKING_DAYS - 1, CLOSE_QUOTE_IDX))
        
    End If
    
    open_quote_val = CDbl(security_data(1, 3)) 'Get first business day opening price
    
    last_business_day_close_quote = CDbl(security_data(record_count, CLOSE_QUOTE_IDX))
    
    If open_quote_val = 0# Then
    
        open_quote_val = CDbl(security_data(2, OPEN_QUOTE_IDX)) 'If the first business day opening price is 0.0, then go with second business day
            
    End If
    
    yearly_change = last_business_day_close_quote - open_quote_val
    
    If open_quote_val > 0 Then
    
        percent_change = (last_business_day_close_quote - open_quote_val) / open_quote_val
            
    End If
        
    
    '(2) Write Data------------------------------------------------------------------
    
    Cells(analysis_row_idx, analysis_col_idx).Select
    
    ticker_val = security_data(1, 1) 'Ticker
    
    ActiveCell.Value = ticker_val
    
    ActiveCell.Offset(0, 1).Select
    
    ActiveCell.Value = yearly_change
    ApplyConditionalFormat
        
    ActiveCell.Offset(0, 1).Select
    
    If open_quote_val > 0 Then
    
        ActiveCell.Value = percent_change
        ActiveCell.NumberFormat = "0.00%"
        
    Else
    
        ActiveCell.Value = ""
        
    End If

    ActiveCell.Offset(0, 1).Value = total_trade_volume
    
    '(3) Keep running track of summary analysis values----------------------------------------
    If open_quote_val > 0 Then
    
        If Str(greatest_pct_increase(2)) = "" Or greatest_pct_increase(2) < percent_change Then
           greatest_pct_increase(1) = ticker_val
           greatest_pct_increase(2) = percent_change
        End If
        
        If Str(greatest_pct_decrease(2)) = "" Or greatest_pct_decrease(2) > percent_change Then
           greatest_pct_decrease(1) = ticker_val
           greatest_pct_decrease(2) = percent_change
        End If
        
        If Str(greatest_tot_vol(2)) = "" Or greatest_tot_vol(2) < total_trade_volume Then
           greatest_tot_vol(1) = ticker_val
           greatest_tot_vol(2) = total_trade_volume
        End If
    
    End If
    
    'Prep for next re-entry in analysis summary
    analysis_row_idx = analysis_row_idx + 1
    
    'Reset variable state---------------------------------
    Range(trades_placeholder).Select
    
End Function

'Description: Applies appropriate cell shading based on value
'Parameters: None
'Returns: Function does not return a value
Private Function ApplyConditionalFormat()

 Const Red As Integer = 3
 Const Green As Integer = 4

    With ActiveCell
    
        .NumberFormat = "0.00"
    
        If .Value < 0# Then
            .Interior.ColorIndex = Red
        ElseIf .Value > 0# Then
            .Interior.ColorIndex = Green
        End If
        
    End With
        
End Function

'Description: Returns the number of equity security records in one trade batch
'Parameters: ByRef security_data() As String - The security trade records to be counted
'Returns: The number of trade records in one and only one trade batch
Private Function GetRecordCount(ByRef security_data() As String) As Integer

    Dim idx As Integer
    Dim count As Integer
    
    count = 0

    For idx = 1 To NUMBER_WORKING_DAYS
    
        If (security_data(idx, 1) <> "") Then
            count = count + 1
        Else
            Exit For
        End If
        
    Next idx
    
    GetRecordCount = count

End Function

'Description: Writes the analysis summary for the current worksheet
'Parameters: None
'Returns: None
Private Function WriteAnalysisSummary()

    Dim idx As Integer
    Range(analysis_summary_col_target).Select
    Dim is_even_idx As Boolean
    
    For idx = 1 To 2
    
        is_even_idx = idx Mod 2 = 0
    
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = greatest_pct_increase(idx)
        
        If is_even_idx Then
            ActiveCell.NumberFormat = "0.00%"
        End If
        
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = greatest_pct_decrease(idx)
        
        If is_even_idx Then
            ActiveCell.NumberFormat = "0.00%"
        End If
        
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = greatest_tot_vol(idx)

        If is_even_idx Then
             ActiveCell.ColumnWidth = Len(ActiveCell.Value)
        End If

        Range(analysis_summary_col_target).Select
        ActiveCell.Offset(0, 1).Select
        
    Next idx

End Function

