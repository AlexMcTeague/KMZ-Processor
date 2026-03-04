Attribute VB_Name = "GeneralUse"
Function Button_AdjacentCell(Optional dir As String = "R") As Range

    Dim cellTarget As Range
    Dim result As Range
    
    Set cellTarget = ActiveSheet.Buttons(Application.Caller).TopLeftCell
    
    Select Case LCase(Left(dir, 1))
        Case "l"
            Set result = cellTarget.Offset(0, -1)
        Case "r"
            Set result = cellTarget.Offset(0, 1)
        Case "u"
            Set result = cellTarget.Offset(-1, 0)
        Case "d"
            Set result = cellTarget.Offset(1, 0)
        Case Else
            MsgBox ("Unrecognized button argument. Ask the maintainer of this spreadsheet for assistance.")
            Set result = cellTarget.Offset(0, 1)
    End Select
    
    Set Button_AdjacentCell = result

End Function

Sub Button_CopyToClipboard(Optional dir As String = "R")

    Dim objCP As Object
    Dim cellTarget As Range
    
    Set objCP = CreateObject("HtmlFile")
    Set cellTarget = Button_AdjacentCell(dir)
    
    objCP.ParentWindow.ClipboardData.SetData "text", CStr(cellTarget.Value)
    
End Sub

Sub Button_SelectFilepath(Optional dir As String = "R")

    Dim cellTarget As Range
    
    Set cellTarget = Button_AdjacentCell(dir)
    
    cellTarget.Value = Application.GetOpenFilename(fileFilter:="Excel Files (*.*), *.*", title:="Select A File")

End Sub

Sub Button_OpenFilepath(Optional dir As String = "R")

    Dim filePath As String
    
    filePath = CStr(Button_AdjacentCell(dir).Value)
    ThisWorkbook.FollowHyperlink filePath

End Sub

Public Function OpenPath(cell As Range) As Workbook
    Dim wb As Workbook

    ' If there isn't a path available, end all macros and inform the user
    path = cell.Value
    If IsEmpty(cell) Or path = "" Then
        ThisWorkbook.Sheets("File Imports").Activate
        cell.Select
        MsgBox (cell.Name.Name & " is not set. Please select a file, then try again.")
        End
    Else
        ' If the file can't be found, end all macros and inform the user
        If dir(path) = "" Then
            ThisWorkbook.Sheets("File Imports").Activate
            cell.Select
            MsgBox ("File doesn't exist at " & path & ". Please select a different file, then try again.")
            End
        End If
        Set wb = Workbooks.Open(path)
        Set OpenPath = wb
    End If
End Function

Public Function SelectFolder() As String
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    Dim FldrPicker As FileDialog
    Dim myFolder As String
    
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    'Have user select folder with Dialog Box
    With FldrPicker
        .title = "Select Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function 'Check if user clicked cancel button
        myFolder = .SelectedItems(1) & "\"
    End With
    SelectFolder = myFolder 'Returns the path to the folder as a String
End Function

Function MakeFolder(ByRef xmlDoc As MSXML2.DOMDocument60, ByVal Name As String) As IXMLDOMNode
    ' Create and return a new named folder node
    Dim newFolder As IXMLDOMNode
    Set newFolder = xmlDoc.createElement("Folder")
    newFolder.appendChild(xmlDoc.createElement("name")).Text = Name
    
    Set MakeFolder = newFolder
End Function

Sub CreateFolderPath(folderPath As String)
    'Create all the folders in a folder path
    'SOURCE: https://exceloffthegrid.com/vba-code-create-delete-manage-folders/
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim arrayElement As Variant

    'Split the folder path into individual folder names
    individualFolders = Split(folderPath, "\")

    'Loop though each individual folder name
    For Each arrayElement In individualFolders

        'Build string of folder path
        tempFolderPath = tempFolderPath & arrayElement & "\"
 
        'If folder does not exist, then create it
        If dir(tempFolderPath, vbDirectory) = "" Then
 
            MkDir tempFolderPath
 
        End If
 
    Next arrayElement
    'End CreateFolders
End Sub



Function ExtractBetween(mainText As String, startString As String, endString As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempString As String

    ' Find the position of the start string
    startPos = InStr(mainText, startString)

    If startPos <> 0 Then
        ' Adjust start position to the character immediately after the start string
        startPos = startPos + Len(startString)

        ' Find the position of the end string, starting the search after the start position
        endPos = InStr(startPos, mainText, endString)

        If endPos <> 0 Then
            ' Extract the substring using Mid
            ' Length is the difference between end and start positions
            ExtractBetween = Mid(mainText, startPos, endPos - startPos)
        Else
            ' Handle case where end string is not found (optional error handling)
            ExtractBetween = ""
        End If
    Else
        ' Handle case where start string is not found (optional error handling)
        ExtractBetween = ""
    End If
End Function

