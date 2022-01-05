Attribute VB_Name = "Handy_Functions"
'
' -- Contents --
'
'    -- Working With Columns --
'        Col_Let_To_Num     - Converts column letters in to their equivalent number
'        Col_Num_To_Let     - Converts a column number in to its equivalent letter(s)
'
'    -- Working With Strings --
'        Chars_Are_In       - Checks all the characters in a string is in a second string
'
'
'
'
Function Col_Let_To_Num(col_let As String) As Double
    
    ' +-------------------------------------------------------------------------------------------+
    ' | -- Purpose --                                                                             |
    ' |    Converts a column letter to its equivalent number                                      |
    ' |                                                                                           |
    ' | -- Arguments --                                                                           |
    ' |    col_let       - The letter(s) to be converted into a number. Input as a string         |
    ' |                                                                                           |
    ' | -- Returned Value --                                                                      |
    ' |                  - Returns the column number if found                                     |
    ' |                  - Returns 0 if the number isnt found (as an error capture)               |
    ' |                                                                                           |
    ' | -- Requirements --                                                                        |
    ' |                  - Needs the Chars_Are_In function available in order to work             |
    ' +-------------------------------------------------------------------------------------------+
    
    ' >>>> Set up variables <<<<
    Dim column_number As Double                                                                                         ' Holds the calculated column number
    
    column_number = 0                                                                                                   ' Set to 0 as a default
    col_let = Trim(UCase(col_let))                                                                                      ' Convert passed string to upper case and trim
    
    ' >>>> Check safe to run <<<<
    If col_let = "" Then GoTo Finished                                                                                  ' If there is no string to work on then finish
    If Not Chars_Are_In(col_let, "ABCDEFGHIJKLMNOPQRSTUVWXYZ") Then GoTo Finished                                       ' Check there are no invalid characters
    
    ' >>>> Convert the letter(s) to a number <<<<
    column_number = Range(col_let & "1").Column                                                                         ' Find the column number
    
    ' >>>> Finish up <<<<
Finished:
    Col_Let_To_Num = column_number                                                                                      ' Set the resulted to the calculated figure

End Function

Function Col_Num_To_Let(col_num As Double) As String
    
    ' +-------------------------------------------------------------------------------------------+
    ' | -- Purpose --                                                                             |
    ' |    Converts a column number to its equivalent letter                                      |
    ' |                                                                                           |
    ' | -- Arguments --                                                                           |
    ' |    col_num       - The number to be converted into a column letter(s). Input as a double  |
    ' |                                                                                           |
    ' | -- Returned Value --                                                                      |
    ' |                  - Returns a string containing the leter(s) if found                      |
    ' |                  - Returns an empty string if there is any kind of error                  |
    ' +-------------------------------------------------------------------------------------------+

    ' >>>> Set up variables <<<<
    Dim column_letter As String                                                                                         ' Holds the calculated letter(s)
    
    column_letter = ""                                                                                                  ' Default the calculated letter(s) to a blank string

    ' >>>> Check safe to run <<<<
    If col_num < 0 Or col_num > 16384 Then GoTo Finished                                                                ' Number is outside the permited number of columns so exit
    If Int(col_num) <> col_num Then GoTo Finished                                                                       ' Column number is not a whole number so exit

    ' >>>> Convert number to letter(s) <<<<
    column_letter = Split(Cells(1, col_num).Address, "$")(1)                                                            ' Find the column letters
    
    ' >>>> Finish up <<<<
Finished:
    Col_Num_To_Let = column_letter

End Function

Function Chars_Are_In(first_string, second_string) As Boolean

    ' +-------------------------------------------------------------------------------------------+
    ' | -- Purpose --                                                                             |
    ' |    Checks all the characters used in the first string are in the second string.           |
    ' |    This is to ensure only valid characters are used. Returns TRUE if valid, FALSE if not. |
    ' |    Process is case sensitive                                                              |
    ' |                                                                                           |
    ' | -- Arguments --                                                                           |
    ' |    first_string  - Contains the string of characters to be checked                        |
    ' |    second_string - Contains the string of characters allowed                              |
    ' |                                                                                           |
    ' | -- Returned Value --                                                                      |
    ' |                  - Returns FALSE by default (including for errors)                        |
    ' |                  - Returns TRUE if all the characters in the first string are found in    |
    ' |                    in the second string. (Is case sensitive)                              |
    ' +-------------------------------------------------------------------------------------------+
    
    ' >>>> Set up variables <<<<
    Dim result As Boolean                                                                                               ' Holds if the values are valid or not
    
    result = False                                                                                                      ' Set result to FALSE as a default
    
    ' >>>> Check safe to run <<<<
    If first_string = "" Or second_string = "" Then GoTo Finished

    ' >>>> Check the strings <<<<
    For f = 1 To Len(first_string)                                                                                      ' Iterate over first string characters
        For s = 1 To Len(second_string)                                                                                 ' Iterate over second string characters
            If Mid(first_string, f, 1) = Mid(second_string, s, 1) Then GoTo Next_First                                  ' If the f char in the first string matches
                                                                                                                        ' char s of the second string move to the next
                                                                                                                        ' char in the first string
        Next s
        GoTo NotFound                                                                                                   ' If we reach this point char f of the first string
                                                                                                                        ' wasnt found in the second string so break out
Next_First:
    Next f
    result = True                                                                                                       ' If we reach point no missing matches were found
                                                                                                                        ' and we can change the result to TRUE
NotFound:
    
    ' >>>> Finish up <<<<
Finished:
    Chars_Are_In = result                                                                                               ' Pass the results back

End Function
