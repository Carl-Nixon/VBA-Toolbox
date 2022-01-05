Attribute VB_Name = "File_Handling"
'
' -- Contents --
'
'    -- Working With Folders --
'        Browse_For_Folder  - Allows the user to browse for a folder
'        Folder_Exist       - Detects if a folder exists
'
'    -- Working With Files --
'        Browse_For_File    - Allows the user to browse for a file
'        File_Exists        - Detects if a file exists
'
'
'
Function Browse_For_Folder(existing_path As String, title As String) As String

    ' +--------------------------------------------------------------------------------------------------+
    ' | -- Purpose --                                                                                    |
    ' |    Allows the user to browse for a folder, starting at either the provided folder or windows     |
    ' |    current active folder. If no folder is found or browsed for, the original starting folder     |
    ' |    is returned instead.                                                                          |
    ' |                                                                                                  |
    ' | -- Arguments --                                                                                  |
    ' |    existing_path - Holds the path to the folder to start searching for.                          |
    ' |                  - If passed a file path including a filename, the filename will be ignored.     |
    ' |                  - If passed a blank string the browsing will start at the last folder accessed  |
    ' |                    in file explorer.                                                             |
    ' |    title         - Holds the title to be shown in the folder browser dialog                      |
    ' |                                                                                                  |
    ' | -- Returned Value --                                                                             |
    ' |                  - Returns the path of the folder browsed for.                                   |
    ' |                  - If no folder is browsed for, the existing_path will be returned               |
    ' +--------------------------------------------------------------------------------------------------+
    
    ' >>>> Set up variables <<<<
    Dim file_dialog As FileDialog                                                                                       ' Holds an instance for the folder dialog window
    Dim dialog_return As String                                                                                         ' Holds the value returnded by the folder dialog window
    Set file_dialog = Application.FileDialog(msoFileDialogFolderPicker)                                                 ' Set the folder dialog window
    
    ' >>>> Prepare existing_path <<<<
    If InStr(1, existing_path, "\") = 0 Then existing_path = ""                                                         ' If there isnt a '\' in the path the path
                                                                                                                        ' isnt valid so set the path to ""
    If Len(existing_path) > 0 Then                                                                                      ' Make sure the file path isnt empty
        If Right(existing_path, 1) <> "\" Then existing_path = Left(existing_path, InStrRev(existing_path, "\") - 1)    ' If exisiting_path is a full file path strip
                                                                                                                        ' file name
    End If
    If Right(existing_path, 1) <> "\" Then existing_path = existing_path & "\"                                          ' Make sure folder path ends in a '\'
    
    ' >>>> Set up and open the file dialog window <<<<
    If Trim(title) <> "" Then                                                                                           ' Was a title provided
        file_dialog.title = title                                                                                       ' Add the title
    End If
    With file_dialog                                                                                                    ' Display dialog
        .title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = existing_path
        If .Show <> -1 Then GoTo NotFound                                                                               ' Breakout if no file found
        dialog_return = .SelectedItems(1)                                                                               ' return the selected item
    End With
NotFound:

    ' >>>> Tidy up the results <<<<
    If dialog_return = "" Then dialog_return = existing_path                                                            ' If nothing was found, then revert to the original path
    If Len(dialog_return) > 0 Then                                                                                      ' Make sure the result isnt empty
        If Right(dialog_return, 1) <> "\" Then dialog_return = dialog_return & "\"                                      ' If the result doesnt end in a '\' add a '\'
    End If
    
    ' >>>> Return the value found/calculated <<<<
    If dialog_return = "\" Then dialog_return = ""                                                                      ' Make sure empty returns are really empty
    Browse_For_Folder = dialog_return                                                                                   ' Return the result
    Set file_dialog = Nothing                                                                                           ' Clear up the memory

End Function

Function Folder_Exist(folder_path As String) As Boolean
    
    ' +--------------------------------------------------------------------------------------------------+
    ' | -- Purpose --                                                                                    |
    ' |    Checks to see if a given folder path exists or not. If it exists TRUE is returned, if not     |
    ' |    FALSE is returned                                                                             |
    ' |                                                                                                  |
    ' | -- Arguments --                                                                                  |
    ' |    folder_path   - Holds the path to the folder to check.                                        |
    ' |                  - If passed a file path including a filename, the filename will be ignored.     |
    ' |                                                                                                  |
    ' | -- Returned Value --                                                                             |
    ' |                  - Returns FALSE by default                                                      |
    ' |                  - Returns TRUE only if the given folder exists                                  |
    ' +--------------------------------------------------------------------------------------------------+
    
    ' >>>> Set up variables <<<<
    Dim result As Boolean                                                                                               ' Holds if the folder is found or not
    Dim found_folder As String                                                                                          ' Holds the name of the folder found when searching
    
    result = False                                                                                                      ' Set result to FALSE as a default
    
    ' >>>> Check its safe to continue <<<<
    If Trim(folder_path) = "" Then GoTo Finished                                                                        ' No folder path given so exit
    
    ' >>>> Check file exists <<<<
    found_folder = Dir(folder_path, vbDirectory)                                                                        ' Search for the file using DIR
    If found_folder <> "" Then result = True                                                                            ' If a file was found set the result to true

    ' >>>> Finish up <<<<
Finished:
    Folder_Exist = result                                                                                               ' Pass back results
    
End Function

Function Browse_For_File(start_path As String, file_type As String, title As String) As String
    
    ' +--------------------------------------------------------------------------------------------------+
    ' | -- Purpose --                                                                                    |
    ' |    Allows the user to browse for a file, starting at either the provided folder or windows       |
    ' |    current active folder. If no file is found or browsed for, an empty string is returned.       |
    ' |                                                                                                  |
    ' | -- Arguments --                                                                                  |
    ' |    start_path    - Holds the path to the folder to start searching in.                           |
    ' |                  - If passed a file path including a filename, the filename will be ignored.     |
    ' |                  - If passed a blank string the browsing will start at the last folder accessed  |
    ' |                    in file explorer.                                                             |
    ' |    file_type     - Holds what type of file to search for. Accepts any file entension             |
    ' |    title         - Holds what the browse window title should be                                  |
    ' |                                                                                                  |
    ' | -- Returned Value --                                                                             |
    ' |                  - Returns the path and name of the file browsed for.                            |
    ' |                  - If no file is browsed for, an empty string is returned                        |
    ' +--------------------------------------------------------------------------------------------------+
    
     ' >>>> Set up variables <<<<
    Dim file_dialog As FileDialog                                                                                       ' Holds an instance for the folder dialog window
    Dim dialog_return As String                                                                                         ' Holds the value returnded by the file dialog window
    
    Set file_dialog = Application.FileDialog(msoFileDialogFilePicker)                                                   ' Set the folder dialog window
    file_dialog.AllowMultiSelect = False                                                                                ' Switch off the ability to select multiple files
    If Trim(file_type) <> "" Then                                                                                       ' Was a file type set
        file_type = "*." & LCase(file_type)                                                                             ' Build the full filter
    End If
    
    ' >>>> Prepare start_path <<<<
    If InStr(1, start_path, "\") = 0 Then existing_path = ""                                                            ' If there isnt a '\' in the path the path
                                                                                                                        ' isnt valid so set the path to ""
    If Len(start_path) > 0 Then                                                                                         ' Make sure the file path isnt empty
        temp_str = Right(start_path, 5)                                                                                 ' Strip the end of the given path
        If InStr(1, temp_str, ".", vbTextCompare) > 0 Then                                                              ' Use this to detect a file extension
            start_path = Left(start_path, InStrRev(existing_path, "\") - 1)                                             ' Strips the file name from any given path
        End If
        If Right(start_path, 1) <> "\" Then start_path = start_path & "\"                                               ' Add the last "\"
    End If

    ' >>>> Browse for file <<<
    file_dialog.InitialFileName = start_path                                                                            ' OPen file dialog at the given folder
    If title <> "" Then                                                                                                 ' Was a title provided
        file_dialog.title = title                                                                                       ' Add the title
    End If
    With file_dialog
        .Filters.Clear                                                                                                  ' Clear previous dialog filters
        .Filters.Add file_type, file_type                                                                               ' Add the filter
        If .Show = -1 Then                                                                                              ' Was an item found
            dialog_return = .SelectedItems.Item(1)                                                                      ' Set the found item
        Else
            dialog_return = ""                                                                                          ' Set the found item to blank
        End If
    End With
    
    ' >>>> Finish up <<<
    Browse_For_File = dialog_return                                                                                     ' Return the results
    
End Function

Function File_Exist(file_path As String) As Boolean
    
    ' +--------------------------------------------------------------------------------------------------+
    ' | -- Purpose --                                                                                    |
    ' |    Checks to see if a given file exists or not. If it exists TRUE is returned, if not FALSE is   |
    ' |    returned.                                                                                     |
    ' |                                                                                                  |
    ' | -- Arguments --                                                                                  |
    ' |    file_path     - Holds the path to the file to check.                                          |
    ' |                                                                                                  |
    ' | -- Returned Value --                                                                             |
    ' |                  - Returns FALSE by default                                                      |
    ' |                  - Returns TRUE only if the given file exists                                  |
    ' +--------------------------------------------------------------------------------------------------+
    
    ' >>>> Set up variables <<<<
    Dim result As Boolean                                                                                               ' Holds if the file is found or not
    Dim found_file As String                                                                                            ' Holds the name of the file found when looking for the file
    
    result = False                                                                                                      ' Set result to FALSE as a default
    
    ' >>>> Check its safe to continue <<<<
    If Trim(file_path) = "" Then GoTo Finished                                                                          ' No folder path given so exit
    
    ' >>>> Check file exists <<<<
    found_file = Dir(file_path)                                                                                         ' Search for the file using DIR
    If found_file <> "" Then result = True                                                                              ' If a file was found set the result to true
    
    ' >>>> Finish up <<<<
Finished:
    File_Exist = result                                                                                                 ' Pass back results
    
End Function
