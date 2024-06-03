Attribute VB_Name = "Loss_Summary_Automation"
Option Explicit
' First, ensure you have the Microsoft VBScript Regular Expressions reference enabled
' Open the VBA editor with Alt + F11.
' Go to Tools > References.
' Check Microsoft VBScript Regular Expressions 5.5
'        ToDo:
'        Get normalized score of coverages to determine order ---> zi = (xi-min(x)) / (max(x)-min(x))
'        Output and format Summary
'        Output and format details
'            Account for page breaks
'        Output and format open claims
'        Deliver standalone workbook

Sub CreateLossSummary()
    Dim all_claims As Scripting.Dictionary
    Dim i As Long
    Dim aggregation, total As Variant
    Dim cpy As Scripting.Dictionary
    Dim all_totals As Scripting.Dictionary
    
    Application.ScreenUpdating = False
    
    ' Call the function to get the dictionary, all_claims
    Set all_claims = ExtractTableToDictionary()
    
    ' Call the function to aggregate claims data using the existing all_claims dictionary
    Set all_totals = AggregateClaimsData(all_claims)
    
    ' Print the items in the dictionary to test
    For Each aggregation In all_totals.Keys
        Set cpy = all_totals.Item(aggregation)
        Debug.Print aggregation & ":"
        For Each total In cpy.Keys
            Debug.Print "    " & total & ": " & cpy(total)
        Next total
    Next aggregation
End Sub
Function ExtractTableToDictionary() As Scripting.Dictionary
    Dim ws As Worksheet
    Dim claim_table As ListObject
    Dim headers() As String
    Dim all_claims As Object, claim_data As Object
    Dim row_number As Long, column_number As Long
    Dim header As String, cell_value As String, primary_key As String
    
    ' Define the worksheet and table
    Set ws = ThisWorkbook.Worksheets("loss_sheet")
    Set claim_table = ws.ListObjects("loss_table")
    
    ' Create dictionaries to store data
    Set all_claims = CreateObject("Scripting.Dictionary")
    
    ' Loop through each row in the table
    For row_number = 1 To claim_table.ListRows.Count
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
        
        'create a composite key to track slowly changing dimensions (claim valuations over time)
        primary_key = claim_data("claim_number") & "_vao_" & claim_data("valuation_date")
        ' Add the row dictionary to the main dictionary
        all_claims.Add primary_key, claim_data
    Next row_number
    
    Set ExtractTableToDictionary = all_claims 'Return the  dictionary
End Function

Function AggregateClaimsData(ByRef input_claims As Scripting.Dictionary) As Scripting.Dictionary
    Dim all_totals As Scripting.Dictionary
    Dim primary_key As Variant
    Dim claim_data As Scripting.Dictionary
    Dim coverage As String, policy_year As String, status As String
    Dim coverage_policy_year As String
    Dim claim_count As Long
    Dim paid_total As Double, reserve_total As Double, incurred_total As Double
    
    ' Create the dictionary to store the aggregated data
    Set all_totals = CreateObject("Scripting.Dictionary")
    
    ' Loop through each claim in the input dictionary
    For Each primary_key In input_claims.Keys
        Set claim_data = input_claims(primary_key)
        
        ' Retrieve necessary values
        coverage = claim_data("coverage")
        policy_year = CInt(claim_data("policy_year"))
        paid_total = CDbl(claim_data("paid"))
        reserve_total = CDbl(claim_data("reserve"))
        incurred_total = CDbl(claim_data("incurred"))
        
        ' Create composite key
        If coverage = "Products Liability" Then
            coverage_policy_year = "General Liability" & "_" & policy_year
        Else
            coverage_policy_year = coverage & "_" & policy_year
        End If
        
        ' Check if the composite key already exists in the aggregated dictionary
        If Not all_totals.Exists(coverage_policy_year) Then
            ' Initialize the nested dictionary for the new composite key
            Set all_totals(coverage_policy_year) = CreateObject("Scripting.Dictionary")
            all_totals(coverage_policy_year)("coverage") = coverage
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
    Set AggregateClaimsData = all_totals
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


