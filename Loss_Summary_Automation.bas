Attribute VB_Name = "Loss_Summary_Automation"
Option Explicit
' First, ensure you have the Microsoft VBScript Regular Expressions reference enabled
' Open the VBA editor with Alt + F11.
' Go to Tools > References.
' Check Microsoft VBScript Regular Expressions 5.5
'        ToDo:
'        Output and format Summary
'        Output and format details
'            Account for page breaks
'        Output and format open claims
'        Deliver standalone workbook

Sub CreateLossSummary()
'    Dim i As Variant, j As Variant
    Dim ws As Worksheet
    Dim all_claims As Scripting.Dictionary
    Dim all_totals_by_year As Scripting.Dictionary
    Dim coverage_totals As Scripting.Dictionary
    Dim coverage_order As Variant
    Dim row As Long, col As Long
    Dim key As Variant
    Dim item As Variant
    Dim subkey As Variant
    Dim coverage As Variant
    Dim startRow As Long
    Dim start_year As Long
    Dim end_year As Long
    Dim summary_headers As Variant
    Dim policy_year As Long
    
    Application.ScreenUpdating = False
    
    ' Call the function to get the dictionary, all_claims
    Set all_claims = ExtractTableToDictionary()
    
    ' Call the function to aggregate claims data using the existing all_claims dictionary
    Set all_totals_by_year = AggregateClaimsByYear(all_claims)
    
    ' Call the function to get totals by coverage for all years combined
    Set coverage_totals = GetCoverageTotals(all_totals_by_year)
    
    ' Call the function to get coverages ordered by importance
    coverage_order = OrderedCoverageKeys(coverage_totals)
    
    ' Create a new worksheet named "Summary"
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Summary")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Summary"
    End If
    On Error GoTo 0

    ' Set initial row for output
    startRow = 4
    row = startRow
    
    ' Loop through the coverage order
    For Each coverage In coverage_order
        ' Add coverage title
        FormatAsCoverageTitle ws.Cells(row, 2), coverage
        row = row + 1

        ' Add column headers and format
        summary_headers = Array("Effective Date", "Expiration Date", "Carrier", "Valuation Date", "Total Claims", "Total Paid", "Total Reserved", "Total Incurred")
        FormatAsColumnHeader ws, row, 2, summary_headers
        row = row + 1
        
        ' Find most recent policy year for this coverage
        start_year = GetMostRecentPolicyYear(all_totals_by_year, coverage)
        end_year = start_year - 5
        
        For policy_year = start_year To end_year Step -1
            row = FormatCoverageSummaryRow(all_totals_by_year, policy_year, row, coverage, ws)
        Next policy_year
    Next coverage
    row = row + 1
    SetSummaryPageFormats ws, row

End Sub

Function FormatCoverageSummaryRow(ByRef all_totals_by_year As Scripting.Dictionary, policy_year As Long, row As Long, coverage As Variant, ws) As Long
    Dim key As Variant
    Dim zero_year As Boolean
    
    zero_year = False
    
    ' Loop through the dictionary to output data for the current coverage
    For Each key In all_totals_by_year.Keys
        If all_totals_by_year(key)("coverage") = coverage And all_totals_by_year(key)("policy_year") = policy_year Then
            With ws.Cells(row, 2)
                .Value = all_totals_by_year(key)("anniversary") & "/" & CStr(all_totals_by_year(key)("policy_year"))
                .NumberFormat = "m/d/yyyy"
            End With
            With ws.Cells(row, 3)
                .Value = all_totals_by_year(key)("anniversary") & "/" & CStr(all_totals_by_year(key)("policy_year") + 1)
                .NumberFormat = "m/d/yyyy"
            End With
            With ws.Cells(row, 4)
                .Value = all_totals_by_year(key)("carrier")
                .NumberFormat = "@"
            End With
            With ws.Cells(row, 5)
                .Value = all_totals_by_year(key)("valuation_date")
                .NumberFormat = "m/d/yyyy"
            End With
            With ws.Cells(row, 6)
                .Value = all_totals_by_year(key)("claim_count")
                .NumberFormat = "#,##0"
            End With
            With ws.Cells(row, 7)
                .Value = all_totals_by_year(key)("paid_total")
                .NumberFormat = "$#,##0"
            End With
            With ws.Cells(row, 8)
                .Value = all_totals_by_year(key)("reserve_total")
                .NumberFormat = "$#,##0"
                If .Value <> 0 Then
                    .Interior.Color = 205
                    .Font.Color = 15921387
                    .Font.Bold = True
                End If
            End With
            With ws.Cells(row, 9)
                .Value = all_totals_by_year(key)("incurred_total")
                .NumberFormat = "$#,##0"
            End With
            
            ' Add borders and center all cells
            AllBorders ws.Range(ws.Cells(row, 2), ws.Cells(row, 9))
            AlignCenter ws.Range(ws.Cells(row, 1), ws.Cells(row, 10))
            row = row + 1
        End If
    Next key

    FormatCoverageSummaryRow = row
End Function

Function ExtractTableToDictionary() As Scripting.Dictionary
    Dim ws As Worksheet
    Dim claim_table As ListObject
    Dim all_claims As Scripting.Dictionary, claim_data As Scripting.Dictionary
    Dim row_number As Long, column_number As Long
    Dim header As String, cell_value As String, primary_key As String
    Dim current_year As Long, policy_year As Long
    Dim policy_year_col_index As Long, claim_number_col_index As Long
    Dim valuation_date_col_index As Long
    Dim claim_number As String, valuation_date As Date
    Dim claim_dates As Object

    ' Define the worksheet and table
    Set ws = ThisWorkbook.Worksheets("loss_sheet")
    Set claim_table = ws.ListObjects("loss_table")

    ' Create dictionaries to store data
    Set all_claims = CreateObject("Scripting.Dictionary")
    Set claim_dates = CreateObject("Scripting.Dictionary")

    ' Get the current year
    current_year = Year(Date)

    ' Find the "policy_year" column index
    policy_year_col_index = 0
    For column_number = 1 To claim_table.ListColumns.Count
        If MakeSnakeCase(CStr(claim_table.ListColumns(column_number).Name)) = "policy_year" Then
            policy_year_col_index = column_number
        ElseIf MakeSnakeCase(CStr(claim_table.ListColumns(column_number).Name)) = "claim_number" Then
            claim_number_col_index = column_number
        ElseIf MakeSnakeCase(CStr(claim_table.ListColumns(column_number).Name)) = "valuation_date" Then
            valuation_date_col_index = column_number
        End If
    Next column_number

    ' Loop through each row in the table
    For row_number = 1 To claim_table.ListRows.Count
        ' Check if policy_year is within the current year or the prior five years
        policy_year = CLng(claim_table.DataBodyRange.Cells(row_number, policy_year_col_index).Value)
        If policy_year >= (current_year - 5) And policy_year <= current_year Then
            claim_number = claim_table.DataBodyRange.Cells(row_number, claim_number_col_index).Value
            valuation_date = claim_table.DataBodyRange.Cells(row_number, valuation_date_col_index).Value
            
            ' Check for existing claim and compare valuation_date
            If claim_dates.Exists(claim_number) Then
                If valuation_date >= claim_dates(claim_number) Then
                    ' Skip this row if the existing valuation_date is more recent
                    GoTo NextRow
                End If
            End If
            
            ' Update the most recent valuation_date for the claim_number
            claim_dates(claim_number) = valuation_date
            
            Set claim_data = CreateObject("Scripting.Dictionary")
            ' Loop through each column in the row
            For column_number = 1 To claim_table.ListColumns.Count
                header = claim_table.ListColumns(column_number).Name
                cell_value = claim_table.DataBodyRange.Cells(row_number, column_number).Value
                ' Apply desired functions
                header = MakeSnakeCase(CStr(header))
                If InStr(1, "claimant_name driver_name coverage carrier cause status", header) Then
                    cell_value = MakeProperCase(cell_value)
                ElseIf InStr(header, "description") Then
                    cell_value = MakeSentenceCase(cell_value)
                End If
                claim_data(header) = cell_value
            Next column_number

            ' Create a composite key to track slowly changing dimensions (claim valuations over time)
            primary_key = claim_data("claim_number") & "_vao_" & claim_data("valuation_date")
            ' Add the row dictionary to the main dictionary
            all_claims.Add primary_key, claim_data
        End If
NextRow:
    Next row_number

    Set ExtractTableToDictionary = all_claims 'Return the  dictionary
End Function

Function AggregateClaimsByYear(ByRef input_claims As Scripting.Dictionary) As Scripting.Dictionary
    Dim all_totals As Scripting.Dictionary
    Dim valuation_date As Variant, primary_key As Variant
    Dim claim_data As Scripting.Dictionary
    Dim coverage As String, policy_year As String, status As String
    Dim anniversary As String, carrier As String
    Dim coverage_policy_year As String
    Dim claim_count As Long
    Dim paid_total As Double, reserve_total As Double, incurred_total As Double
    
    ' Create the dictionary to store the aggregated data
    Set all_totals = CreateObject("Scripting.Dictionary")
    
    ' Loop through each claim in the input dictionary
    For Each primary_key In input_claims.Keys
        Set claim_data = input_claims(primary_key)
        ' Retrieve necessary values
        anniversary = claim_data("anniversary")
        carrier = claim_data("carrier")
        coverage = claim_data("coverage")
        policy_year = CInt(claim_data("policy_year"))
        paid_total = CDbl(claim_data("paid"))
        reserve_total = CDbl(claim_data("reserve"))
        incurred_total = CDbl(claim_data("incurred"))
        If valuation_date <= claim_data("valuation_date") Then
            valuation_date = claim_data("valuation_date")
        End If
        
        ' Create composite key
        If coverage = "Products Liability" Then
            coverage_policy_year = "General Liability" & " " & policy_year
            coverage = "General Liability"
        Else
            coverage_policy_year = coverage & " " & policy_year
        End If
        
        ' Check if the composite key already exists in the aggregated dictionary
        If Not all_totals.Exists(coverage_policy_year) Then
            ' Initialize the nested dictionary for the new composite key
            Set all_totals(coverage_policy_year) = CreateObject("Scripting.Dictionary")
            all_totals(coverage_policy_year)("anniversary") = anniversary
            all_totals(coverage_policy_year)("carrier") = carrier
            all_totals(coverage_policy_year)("coverage") = coverage
            all_totals(coverage_policy_year)("valuation_date") = valuation_date
            all_totals(coverage_policy_year)("policy_year") = policy_year
            all_totals(coverage_policy_year)("open_claim_count") = 0
            all_totals(coverage_policy_year)("claim_count") = 0
            all_totals(coverage_policy_year)("paid_total") = 0
            all_totals(coverage_policy_year)("reserve_total") = 0
            all_totals(coverage_policy_year)("incurred_total") = 0
        End If
        
        ' Update the aggregated values
        If claim_data("status") <> "Closed" Then
            all_totals(coverage_policy_year)("open_claim_count") = all_totals(coverage_policy_year)("open_claim_count") + 1
        End If
        all_totals(coverage_policy_year)("claim_count") = all_totals(coverage_policy_year)("claim_count") + 1
        all_totals(coverage_policy_year)("paid_total") = all_totals(coverage_policy_year)("paid_total") + paid_total
        all_totals(coverage_policy_year)("reserve_total") = all_totals(coverage_policy_year)("reserve_total") + reserve_total
        all_totals(coverage_policy_year)("incurred_total") = all_totals(coverage_policy_year)("incurred_total") + incurred_total
    Next primary_key
    
    ' Return the aggregated dictionary
    Set AggregateClaimsByYear = all_totals
End Function

Function GetCoverageTotals(ByRef all_totals As Scripting.Dictionary) As Scripting.Dictionary
    Dim grand_totals As Object
    Set grand_totals = CreateObject("Scripting.Dictionary")
    
    Dim coverage_policy_year As Variant
    For Each coverage_policy_year In all_totals.Keys
        Dim coverage_year_data As Object
        Set coverage_year_data = all_totals(coverage_policy_year)
        
        Dim current_coverage As String
        current_coverage = coverage_year_data("coverage")

        If Not grand_totals.Exists(current_coverage) Then
            Set grand_totals(current_coverage) = CreateObject("Scripting.Dictionary")
            grand_totals(current_coverage)("policy_year") = grand_totals(current_coverage)("policy_year")
            grand_totals(current_coverage)("open_claim_count") = 0
            grand_totals(current_coverage)("claim_count") = 0
            grand_totals(current_coverage)("paid_total") = 0#
            grand_totals(current_coverage)("reserve_total") = 0#
            grand_totals(current_coverage)("incurred_total") = 0#
        End If
        
        grand_totals(current_coverage)("open_claim_count") = grand_totals(current_coverage)("open_claim_count") + coverage_year_data("open_claim_count")
        grand_totals(current_coverage)("claim_count") = grand_totals(current_coverage)("claim_count") + coverage_year_data("claim_count")
        grand_totals(current_coverage)("paid_total") = grand_totals(current_coverage)("paid_total") + coverage_year_data("paid_total")
        grand_totals(current_coverage)("reserve_total") = grand_totals(current_coverage)("reserve_total") + coverage_year_data("reserve_total")
        grand_totals(current_coverage)("incurred_total") = grand_totals(current_coverage)("incurred_total") + coverage_year_data("incurred_total")
    Next coverage_policy_year
    
    Set GetCoverageTotals = grand_totals
End Function

Function OrderedCoverageKeys(ByRef coverage_totals As Scripting.Dictionary) As Variant

    ' Call the function to get minimum and maximum values for each metric
    Dim min_max As Scripting.Dictionary
    Set min_max = GetMinimumAndMaximumValues(coverage_totals)
    
    Dim minimum_values As Scripting.Dictionary
    Dim maximum_values As Scripting.Dictionary
    Set minimum_values = min_max("minimum_values")
    Set maximum_values = min_max("maximum_values")

    ' Call the function to get normalized values for each coverage type
    Dim normalized_values As Scripting.Dictionary
    Set normalized_values = GetNormalizedValues(coverage_totals, minimum_values, maximum_values)
    
    ' Call the function to calculate the adjusted average
    Dim adjusted_averages As Scripting.Dictionary
    Set adjusted_averages = GetAdjustedAverages(coverage_totals, normalized_values)
    
    ' Call the function to sort coverages by adjusted average
    Dim sorted_coverages As Variant
    sorted_coverages = SortCoverages(adjusted_averages)
    
    OrderedCoverageKeys = sorted_coverages
End Function
Function GetMinimumAndMaximumValues(ByRef coverage_totals As Scripting.Dictionary) As Scripting.Dictionary
    Dim minimum_values As Scripting.Dictionary
    Dim maximum_values As Scripting.Dictionary
    Set minimum_values = CreateObject("Scripting.Dictionary")
    Set maximum_values = CreateObject("Scripting.Dictionary")

    Dim coverage_keys As Variant
    coverage_keys = coverage_totals.Keys
    
    Dim key As Variant
    Dim inner_key As Variant
    
    For Each key In coverage_keys
        For Each inner_key In coverage_totals(key).Keys
            If Not minimum_values.Exists(inner_key) Then
                minimum_values(inner_key) = coverage_totals(key)(inner_key)
                maximum_values(inner_key) = coverage_totals(key)(inner_key)
            Else
                If coverage_totals(key)(inner_key) < minimum_values(inner_key) Then
                    minimum_values(inner_key) = coverage_totals(key)(inner_key)
                End If
                If coverage_totals(key)(inner_key) > maximum_values(inner_key) Then
                    maximum_values(inner_key) = coverage_totals(key)(inner_key)
                End If
            End If
        Next inner_key
    Next key
    
    Dim min_max_collection As Scripting.Dictionary
    Set min_max_collection = CreateObject("Scripting.Dictionary")
    min_max_collection.Add "minimum_values", minimum_values
    min_max_collection.Add "maximum_values", maximum_values
    
    Set GetMinimumAndMaximumValues = min_max_collection
End Function

Function GetNormalizedValues(ByRef coverage_totals As Scripting.Dictionary, ByRef minimum_values As Scripting.Dictionary, ByRef maximum_values As Scripting.Dictionary) As Scripting.Dictionary
    Dim normalized_values As Scripting.Dictionary
    Set normalized_values = CreateObject("Scripting.Dictionary")
    
    Dim coverage_keys As Variant
    coverage_keys = coverage_totals.Keys
    
    Dim key As Variant
    Dim inner_key As Variant
    
    For Each key In coverage_keys
        Set normalized_values(key) = CreateObject("Scripting.Dictionary")
        For Each inner_key In coverage_totals(key).Keys
            If inner_key = "policy_year" Then
                GoTo next_inner_key
            End If
            If maximum_values(inner_key) <> minimum_values(inner_key) Then
                normalized_values(key)(inner_key) = (coverage_totals(key)(inner_key) - minimum_values(inner_key)) / (maximum_values(inner_key) - minimum_values(inner_key))
            Else
                normalized_values(key)(inner_key) = 0
            End If
next_inner_key:
        Next inner_key
    Next key
    
    Set GetNormalizedValues = normalized_values
End Function

Function GetAdjustedAverages(ByRef coverage_totals As Scripting.Dictionary, ByRef normalized_values As Scripting.Dictionary) As Scripting.Dictionary
    Dim weighted_averages As Scripting.Dictionary
    Set weighted_averages = CreateObject("Scripting.Dictionary")
    
    Dim coverage_keys As Variant
    coverage_keys = coverage_totals.Keys
    
    Dim key As Variant
    Dim inner_key As Variant
    
    For Each key In coverage_keys
        Dim sum_normalized As Double
        sum_normalized = 0
        
        For Each inner_key In normalized_values(key).Keys
            sum_normalized = sum_normalized + normalized_values(key)(inner_key)
        Next inner_key
        
        Dim weight As Double
        weight = 0.01 * (coverage_totals(key)("open_claim_count") / coverage_totals(key)("claim_count"))
        
        Dim average_normalized As Double
        average_normalized = sum_normalized / coverage_totals(key).Count
        
        Dim weighted_average As Double
        weighted_average = average_normalized + weight
        
        weighted_averages(key) = weighted_average
    Next key
    
    Set GetAdjustedAverages = weighted_averages
End Function

Function SortCoverages(ByRef weighted_averages As Scripting.Dictionary) As Variant
    Dim sorted_coverage_keys As Collection
    Set sorted_coverage_keys = New Collection
    
    Dim coverage_keys As Variant
    coverage_keys = weighted_averages.Keys
    
    Dim key As Variant
    
    For Each key In coverage_keys
        If sorted_coverage_keys.Count = 0 Then
            sorted_coverage_keys.Add key
        Else
            Dim added As Boolean
            added = False
            Dim i As Long
            For i = 1 To sorted_coverage_keys.Count
                If weighted_averages(key) > weighted_averages(sorted_coverage_keys(i)) Then
                    sorted_coverage_keys.Add key, , i
                    added = True
                    Exit For
                End If
            Next i
            If Not added Then
                sorted_coverage_keys.Add key
            End If
        End If
    Next key
        
    ' Convert sorted coverage keys to array of strings
    Dim result() As String
    ReDim result(1 To sorted_coverage_keys.Count)
    
    For i = 1 To sorted_coverage_keys.Count
        result(i) = sorted_coverage_keys(i)
    Next i
    
    SortCoverages = result
End Function

Function MakeSnakeCase(input_string As String)
    Dim regEx As Object
    Dim i As Integer
    
    ' Create a regex object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True

    input_string = LCase(Trim(input_string)) ' Trim the string and convert it to lowercase
    regEx.Pattern = "[^\w\s]" ' Regex pattern for removing all non-word and non-whitespace characters
    input_string = regEx.Replace(input_string, "") ' Remove all non-word and whitespace characters
    regEx.Pattern = "\s+" ' Regex pattern for replacing whitespace with underscores
    input_string = regEx.Replace(input_string, "_") ' Replace spaces within the string with underscores
    
    ' Return the modified string
    MakeSnakeCase = input_string
End Function

Function MakeProperCase(input_string As String)
    MakeProperCase = Application.WorksheetFunction.Proper(input_string)
End Function

Function MakeSentenceCase(input_string As String)
    Dim regEx As Object
    Dim string_item As String
    Dim i As Long ' Loop counter
    Dim new_string As String ' String to hold the result
    Dim new_word, new_character As String ' Current word and character being processed
    Dim is_new_sentence As Boolean ' Flag to indicate the start of a new sentence
    
    ' Create a regex object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    ' Define the regex pattern for removing special characters
    regEx.Pattern = "[@#$%^&*+_={}\[\]\|\\\/<>`~]"
        
    input_string = Trim(input_string) ' Remove leading and trailing whitespace
    input_string = regEx.Replace(input_string, "") ' Remove special characters in regex pattern
    
    If Len(input_string) = 0 Then ' If the input string is empty,
        MakeSentenceCase = ""     ' return an empty string
        Exit Function
    End If
    
    is_new_sentence = False ' Initialize the new sentence flag
    input_string = LCase(input_string) ' Convert the entire string to lowercase
    new_string = UCase(Left(input_string, 1)) ' Capitalize the first character
    For i = 2 To Len(input_string) ' Loop through the rest of the string
        new_character = Mid(input_string, i, 1)

        If is_new_sentence Then
            If new_character >= "a" And new_character <= "z" Then
                ' Capitalize the character if it is the start of a new sentence
                new_string = Trim(new_string) & " " & UCase(new_character)
                is_new_sentence = False  ' Reset the new sentence flag
            End If
        Else
            ' Otherwise, keep the character in lowercase
            new_string = new_string & new_character
        End If
        
        ' Check if the current character is a sentence-ending punctuation
        If new_character = "." Or new_character = "!" Or new_character = "?" Then
            is_new_sentence = True ' Set the new sentence flag
        End If
    Next i
        
    ' Add a period at the end if it does not end with '.', '!', or '?'
    If Right(new_string, 1) <> "." And Right(new_string, 1) <> "!" And Right(new_string, 1) <> "?" Then
        new_string = new_string & "."
    End If

    ' Return the modified string
    MakeSentenceCase = new_string
End Function

Function AlignCenter(rng As Range)
    With rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Function

Function AllBorders(rng As Range)
    With rng
        .Borders.weight = xlThin
    End With
End Function

Function FormatAsCoverageTitle(rng As Range, coverage As Variant)
    With rng
        .Value = coverage
        .Interior.Pattern = xlNone
        .Font.Name = "Aptos"
        .Font.Size = 12
        .Font.Bold = True
        .Font.ColorIndex = xlAutomatic
    End With
End Function

Function FormatAsColumnHeader(ws As Worksheet, row As Long, start_column As Long, title_array As Variant)
    Dim title As Variant
    Dim col As Long
    col = start_column
      
    For Each title In title_array
        With ws.Cells(row, col)
            .Value = title
            .Interior.Color = 4337152
            .Font.Name = "Aptos Narrow"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = 15921387
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        col = col + 1
        
    Next title
End Function

Function SetSummaryPageFormats(ws As Worksheet, row As Long)
    Dim endRow As Long
    ' Apply formatting and page setup
    With ws
        ' Set column widths
        .Columns("A").ColumnWidth = 0.42
        .Columns("J").ColumnWidth = 0.42
        .Columns("B:C").ColumnWidth = 15
        .Columns("D:I").AutoFit
        ' Set row heights
        .Rows(1).RowHeight = 3.75
        .Rows(row - 1).RowHeight = 3.75
        ' Set print format
        .PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(row - 1, 10)).Address
        .PageSetup.PrintTitleRows = "$1:$3"
        .PageSetup.Orientation = xlLandscape
        ' Insert page breaks
        endRow = 35
        While endRow < row
            .Rows(endRow + 1).PageBreak = xlPageBreakManual
            endRow = endRow + 35
        Wend
    End With
End Function

Function GetMostRecentPolicyYear(ByRef all_totals_by_year As Scripting.Dictionary, coverage As Variant) As Long
Dim most_recent_year_seen As Long
Dim current_year As Long
Dim current_month As Long
Dim coverage_policy_year As Variant
Dim most_recent_year As Variant
Dim anniversary_month As Variant

' Get current year and month
current_year = Year(Date)
current_month = Month(Date)

For Each coverage_policy_year In all_totals_by_year.Keys
    If all_totals_by_year(coverage_policy_year)("coverage") = coverage Then
        anniversary_month = Left(all_totals_by_year(coverage_policy_year)("anniversary"), 2)
        If CLng(anniversary_month) > current_month Then
            most_recent_year = current_year - 1
        Else
            most_recent_year = current_year
        End If
    End If
Next coverage_policy_year

GetMostRecentPolicyYear = most_recent_year
End Function
