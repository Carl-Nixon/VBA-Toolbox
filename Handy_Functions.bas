Attribute VB_Name = "Handy_Functions"
' Contents
'   Working With Columns
'       ColLetToNum            - Converts column letters in to their equivalent number
'       ColNumToLet            - Converts a column number in to its equivalent letter(s)
'   Working With Strings
'       CharsOnlyIn            - Checks all the characters in a string is in a second string
'
Function ColLetToNum(col_let As String) As Integer
    
    ' Purpose -->   Converts a column letter to its equivalent number
    ' Arguments --> col_let (string)
    '               - The letter(s) to be converted into a number
    ' Returns -->   (Integer)
    '               - Returns the column number if found
    '               - Returns 0 if the number isn't found (as an error capture)
    
    ' Set up variables
    ' ================
    
    ' Holds the calculated column number
    Dim column_number As Integer
    
    ' Set to 0 as a default
    column_number = 0
    
    ' Convert passed string to upper case and trim it to make it safer
    col_let = Trim(UCase(col_let))
    
    ' Check it is safe to run
    ' =======================
    
    ' If there is no string to work on then finish
    If col_let = "" Then GoTo Finished

    ' Convert the letter(s) to a number
    ' =================================
    
    ' Build a range using the column letter then split it
    ' back out so the column number can be extracted
    column_number = Range(col_let & "1").Column
    
    ' Finish up
    ' =========
Finished:

    ' Return the calculated column number
    ColLetToNum = column_number

End Function

Function ColNumToLet(col_num As Integer) As String
    
    ' Purpose -->   Converts a column number to its equivalent letter
    ' Arguments --> col_num(integer)
    '               - The number to be converted into a column letter(s).
    ' Returns -->   (String)
    '               - Returns a string containing the letter(s) if found
    '               - Returns an empty string if there is any kind of error

    ' Set up variables
    ' ================
    
    ' Holds the calculated letter(s)
    Dim column_letter As String
    
    ' Default the calculated letter(s) to a blank string
    column_letter = ""

    ' Check safe to run
    ' =================
    
    ' If number is outside the permitted number of columns then exit
    If col_num < 0 Or col_num > 16384 Then GoTo Finished
    
    ' If column number is not a whole number then exit
    If Int(col_num) <> col_num Then GoTo Finished

    ' Convert number to letter(s)
    ' ===========================
    
    ' Build a range using the column number then split it
    ' back out so the column letter can be extracted
    column_letter = Split(Cells(1, col_num).Address, "$")(1)
    
    ' Finish up
    ' =========
Finished:

    ' Return the calculated column letter
    ColNumToLet = column_letter

End Function

Function CharsOnlyIn(first_str, second_str) As Boolean

    ' Purpose -->   Checks all the characters used in the first string are in the
    '               second string. This is to ensure only valid characters are used.
    '               Returns TRUE if valid, FALSE if not. Process is case sensitive
    ' Arguments --> first_str (string)
    '               - Contains the string of characters to be checked
    '                 second_str (string)
    '               - Contains the string of characters allowed
    ' Returns -->   (Boolean)
    '               - Returns FALSE by default (including for errors)
    '               - Returns TRUE if all the characters in the first string are
    '                 found in the second string.
    
    ' Set up variables
    ' ================
    
    ' Holds the calculated result to be returned at the end
    Dim result As Boolean
    
    ' Default the result to false
    result = False
    
    ' Check its safe to proceed
    ' =========================
    
    ' If either string is empty then jump to finish as a result cant be calculated
    If first_str = "" Or second_str = "" Then GoTo Finished

    ' Check the strings
    ' =================
    
    ' Iterate over first string characters
    For f = 1 To Len(first_str)
        
        ' Iterate over second string characters
        For s = 1 To Len(second_str)
            
            ' If the f char in the first string matches chars of the
            ' second string move to the next char in the first string
            If Mid(first_str, f, 1) = Mid(second_str, s, 1) Then GoTo Next_First
                                                                                        
        Next s
        
        ' If we reach this point char f of the first string
        ' wasn't found in the second string so break out
        GoTo NotFound
        
Next_First:
    Next f
    
    ' If we reach point no missing matches were found
    ' and we can change the result to TRUE
    result = True
    
NotFound:
    
    ' Finish up
    ' =========
Finished:

    ' Pass the result back
    CharsOnlyIn = result

End Function
