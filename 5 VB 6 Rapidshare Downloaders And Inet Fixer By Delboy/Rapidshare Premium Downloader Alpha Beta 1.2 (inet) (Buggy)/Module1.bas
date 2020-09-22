Attribute VB_Name = "Module1"
Public Sub Loadlistbox(Directory As String, TheList As listbox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub

Public Sub SaveListBox(Directory As String, TheList As listbox)
    
    Dim savelist As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For savelist& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(savelist&)
    Next savelist&
    Close #1
End Sub

Public Sub xListRemoveSelected(listbox As listbox)
        Dim ListCount As Long
ListCount& = listbox.ListCount
Do While ListCount& > 0&
ListCount& = ListCount& - 1
If listbox.Selected(ListCount&) = True Then
listbox.RemoveItem (ListCount&)
End If
Loop
End Sub

Public Function SaveText(Text As String, FileName As String) As Boolean
On Error GoTo handle
Dim sTemp As String
    sTemp = Text
    Open FileName For Append As #1  'Opening the file to SaveText
        Print #1, sTemp             'Printing  the text to the file
    Close #1                        'Closing
    If FileExists(FileName) = False Then    'Check whether the file created
        MsgBox "Unexpectd error occured. File could not be saved", vbCritical, "Sorry"
        SaveText = False    'Returns 'False'
    Else
        SaveText = True     'Returns 'True'
    End If
Exit Function
handle:
    SaveText = False
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Function
Public Function FileExists(FileName As String) As Boolean
'This function checks the existance of a file
On Error GoTo handle
    If FileLen(FileName) >= 0 Then: FileExists = True: Exit Function
handle:
    FileExists = False
End Function
