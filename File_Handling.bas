Attribute VB_Name = "File_Handling"
' Contents
'   Working With Folders --
'       BrowseFolder           - Allows the user to browse for a folder
'       DoesFolderExist        - Detects if a folder exists
'   Working With Files --
'       BrowseForFile          - Allows the user to browse for a file
'       DoesFileExist          - Detects if a file exists
'
Function BrowseFolder(existing_path As String, title As String) As String
    
    ' Purpose -->   Allows the user to browse for a folder, starting at either the provided
    '               folder or windows current active folder. If no folder is found or browsed
    '               for, the original starting folder is returned instead.
    ' Arguments --> existing_path (string)
    '               - Holds the path to the folder to start searching for.
    '               - If passed a file path including a filename, the filename will be ignored.
    '               - If passed a blank string the browsing will start at the last folder
    '                 accessed in file explorer.
    '               title (string)
    '               - Holds the title to be shown in the folder browser dialog
    ' Returns -->   (String)
    '               - Returns the path of the folder browsed for.
    '               - If no folder is browsed for, the existing_path will be returned
    
    ' Set up variables
    ' ================
    
    ' Holds an instance for the folder dialog window
    Dim file_dialog As FileDialog
    
    ' Holds the value returned by the folder picker dialog window
    Dim dialog_return As String
    
    ' Open the folder dialog window
    Set file_dialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' Prepare existing_path
    ' =====================
    
    ' If there isn't a '\' in the path, the path isn't valid so set the path to ""
    If InStr(1, existing_path, "\") = 0 Then existing_path = ""
    
    ' Make sure the file path isnt empty
    If Len(existing_path) > 0 Then
        
        ' If exisiting_path is a full file path (not a folder path)
        If Right(existing_path, 1) <> "\" Then
        
            ' If it is a folder path, remove the file name part
            existing_path = Left(existing_path, InStrRev(existing_path, "\") - 1)
        
        End If
                                                                                        
    End If
    
    ' Make sure folder path ends in a '\'
    If Right(existing_path, 1) <> "\" Then existing_path = existing_path & "\"
    
    ' Set up and open the file dialog window
    ' ======================================
    
    ' Was the dialog title set?
    If Trim(title) <> "" Then
        
        ' Set the dialog title
        file_dialog.title = title
    
    End If
    
    ' Work with the dialog window
    With file_dialog
    
        ' Set the dialog windows settings
        .AllowMultiSelect = False
        .InitialFileName = existing_path
        
        ' Breakout if not file found
        If .Show <> -1 Then GoTo NotFound
        
        ' Grab the selected values
        dialog_return = .SelectedItems(1)
    
    End With
NotFound:

    ' Tidy up the results
    ' ===================
    
    ' If nothing was selected, then revert to the original path
    If dialog_return = "" Then dialog_return = existing_path
    
    ' Make sure the selected path isnt empty
    If Len(dialog_return) > 0 Then
        
        ' If the result doesn't end in a '\' add a '\'
        If Right(dialog_return, 1) <> "\" Then dialog_return = dialog_return & "\"
    
    End If
    
    ' Finish up
    ' =========
    
    ' Make sure empty returns are really empty
    If dialog_return = "\" Then dialog_return = ""
    
    ' Clear up the memory
    Set file_dialog = Nothing
    
    ' Return the result
    BrowseFolder = dialog_return
    
End Function

Function DoesFolderExist(folder_path As String) As Boolean
    
    ' Purpose -->   Checks to see if a given folder path exists or not. If
    '               it exists TRUE is returned, if not FALSE is returned.
    ' Arguments --> folder_path (string)
    '               - Holds the path to the folder to check. If passed a file
    '                 path including a filename, the filename will be ignored.
    ' Returns -->   (Boolean)
    '               - Returns FALSE by default
    '               - Returns TRUE only if the given folder exists
    
    ' Set up variables
    ' ================
    
    ' Holds if the folder is found or not
    Dim result As Boolean
    
    ' Holds the name of the folder found when searching
    Dim found_folder As String
    
    ' Set result to FALSE as a default
    result = False
    
    ' Check its safe to continue
    ' ==========================
    
    ' If there is no folder path then exit
    If Trim(folder_path) = "" Then GoTo Finished
    
    ' Check folder exists
    ' ===================
    
    ' Search for the file using DIR
    found_folder = Dir(folder_path, vbDirectory)
    
    ' If a returned folder was found then set the result to true
    If found_folder <> "" Then result = True

    ' Finish up
    ' =========
Finished:

    ' Return the result
    DoesFolderExist = result
    
End Function

Function BrowseFile(start_path As String, file_type As String, title As String) As String
    
    ' Purpose -->   Allows the user to browse for a file, starting at either the
    '               provided folder or windows current active folder. If no file is
    '               found or browsed for, an empty string is returned.
    ' Arguments --> start_path (string)
    '               - Holds the path to the folder to start searching in.
    '               - If passed a file path including a filename, the filename will be
    '                 ignored.
    '               - If passed a blank string the browsing will start at the last
    '                 folder accessed in file explorer.
    '               file_type (string)
    '               - Holds what type of file to search for. Accepts any file extension
    '                 title (string)
    '               - Holds what the browse window title should be
    ' Returns -->   (String)
    '               - Returns the path and name of the file browsed for.
    '               - If no file is browsed for, an empty string is returned
    
     ' Set up variables
     ' ================
    
    ' Holds an instance for the folder dialog window
    Dim file_dialog As FileDialog
    
    ' Holds the value returned by the file dialog window
    Dim dialog_return As String
    
    ' Set the folder dialog window
    Set file_dialog = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Switch off the ability to select multiple files
    file_dialog.AllowMultiSelect = False
    
    ' Was a file type set
    If Trim(file_type) <> "" Then
        
        ' Build a default file filter
        file_type = "*." & LCase(file_type)
    
    End If
    
    ' Prepare start_path
    ' ==================
    
    ' If there isn't a '\' in the path, the path isn't valid so set the path to ""
    If InStr(1, start_path, "\") = 0 Then existing_path = ""
    
    ' Make sure the file path isnt empty
    If Len(start_path) > 0 Then
        
        ' Put the last 5 characters of the file path in a string
        temp_str = Right(start_path, 5)
        
        ' Is there "." in the temp_str, it it is then there must a file extension
        If InStr(1, temp_str, ".", vbTextCompare) > 0 Then
        
            ' If there is a file extension, we know to remove the file name from the path
            start_path = Left(start_path, InStrRev(start_path, "\") - 1)

        End If
        
        ' If there isnt a "\" at the end of the file path add it back in
        If Right(start_path, 1) <> "\" Then start_path = start_path & "\"
    
    End If

    ' Browse for file
    ' ===============
    
    ' Open the file browser dialog at the given start location
    file_dialog.InitialFileName = start_path
    
    ' Was a title for the file dialog provided
    If title <> "" Then
        
        ' Add the title to the file dialog window
        file_dialog.title = title
    
    End If
    
    ' Work with the dialog window
    With file_dialog
        
        ' Clear any previous filters and add the selected filters
        .Filters.Clear
        .Filters.Add file_type, file_type
        
        ' Set the selected item to a default blank value
        dialog_return = ""
        
        ' If an item was selected put in the dialog_return variable
        If .Show = -1 Then dialog_return = .SelectedItems.Item(1)
                
    End With
    
    ' Finish up
    ' =========
    
    '' Return the result
    BrowseFile = dialog_return
    
End Function

Function DoesFileExist(file_path As String) As Boolean
    
    ' Purpose -->   Checks to see if a given file exists or not. If it
    '               it exists TRUE is returned, if not FALSE is returned.
    ' Arguments --> file_path (string)
    '               - Holds the path to the file to check.
    ' Returns -->   (Boolean)
    '               - Returns FALSE by default
    '               - Returns TRUE only if the given file exists
    
    ' Set up variables
    ' ================
    
    ' Holds if the file is found or not
    Dim result As Boolean
    
    ' Holds the name of the file found when looking for the file
    Dim found_file As String
    
    ' Set result to FALSE as a default
    result = False
    
    ' Check it is safe to continue
    ' ============================
    
    ' No folder path given so exit
    If Trim(file_path) = "" Then GoTo Finished
    
    ' Check if file exists
    ' ====================
    
    ' Search for the file using DIR
    found_file = Dir(file_path)
    
    ' If a file was found set the result to true
    If found_file <> "" Then result = True
    
    ' Finish up
    ' =========
Finished:

    ' Return the result
    DoesFileExist = result
    
End Function
