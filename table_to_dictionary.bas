Attribute VB_Name = "Module1"
Option Explicit
' First, ensure you have the Microsoft VBScript Regular Expressions reference enabled
' Open the VBA editor with Alt + F11.
' Go to Tools > References.
' Check Microsoft VBScript Regular Expressions 5.5

Function ExtractTableToDictionary() As Scripting.Dictionary
    Dim ws As Worksheet
    Dim claim_table As ListObject
    Dim headers() As String
    Dim all_claims, claim_data As Object
    Dim row_number, column_number As Long
    Dim header, cell_value As String
    
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
        ' Add the row dictionary to the main dictionary
        all_claims.Add claim_data("claim_number"&), claim_data
    Next row_number
    
    Set ExtractTableToDictionary = all_claims 'Return the  dictionary
End Function

Sub TestExtractTableToDictionary()
    Dim claimData As Scripting.Dictionary
    Dim i As Long
    Dim claimNumber, key As Variant
    Dim rowDict As Scripting.Dictionary
    
    ' Call the function to get the dictionary
    Set claimData = ExtractTableToDictionary()
    
    ' Print the items in the dictionary
    For Each claimNumber In claimData.Keys
        Set rowDict = claimData.Item(claimNumber)
        Debug.Print "Claim Number " & claimNumber & ":"
        For Each key In rowDict.Keys
            Debug.Print "    " & key & ": " & rowDict(key)
        Next key
    Next claimNumber
End Sub


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
    input_string = regEx.Replace(input_string, "") ' Remove special charactes in regex pattern
    
    ' If the input string is empty, return an empty string
    If Len(input_string) = 0 Then
        MakeSentenceCase = ""
        Exit Function
    End If
        
    input_string = LCase(input_string) ' Convert the entire string to lowercase
    new_string = UCase(Left(input_string, 1)) ' Capitalize the first character
    is_new_sentence = False ' Initialize the new sentence flag
        
    ' Loop through the rest of the string
    For i = 2 To Len(input_string)
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


