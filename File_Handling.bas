Attribute VB_Name = "File_Handling"
Function Browse_For_Folder(existing_path As String) As String

    ' +-------------------------------------------------------------------------------------------------+
    ' | -- Purpose --                                                                                   |
    ' |   Allows the user to browse for a folder, starting at either the provided folder or windows     |
    ' |   current active folder. If no folder is found or browsed for the original starting folder      |
    ' |   is returned instead.                                                                          |
    ' |                                                                                                 |
    ' | -- Arguments --                                                                                 |
    ' |   existing_path - Holds the path to the folder to start searching for.                          |
    ' |                 - If passed a file path including a filename, the filename will be ignored.     |
    ' |                 - If passed a blank string the browsing will start at the last folder accessed  |
    ' |                   in file explorer.                                                             |
    ' |                                                                                                 |
    ' | -- Returned Value --                                                                            |
    ' |                 - Returns the path of the folder browsed for.                                   |
    ' |                 - If no folder is browsed for the existing_path will be returned                |
    ' +-------------------------------------------------------------------------------------------------+
    
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
    With file_dialog
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = existing_path
        If .Show <> -1 Then GoTo NextCode
        dialog_return = .SelectedItems(1)
    End With
NextCode:

    ' >>>> Tidy up the results <<<<
    If dialog_return = "" Then dialog_return = existing_path                                                            ' If nothing was found, then revert to the original path
    If Len(dialog_return) > 0 Then                                                                                      ' Make sure the result isnt empty
        If Right(dialog_return, 1) <> "\" Then dialog_return = dialog_return & "\"                                      ' If the result doesnt end in a '\' add a '\'
    End If
    
    ' >>>> Return the value found/calculated <<<<
    Browse_For_Folder = dialog_return                                                                                   ' Return the result
    Set file_dialog = Nothing                                                                                           ' Clear up the memory

End Function
