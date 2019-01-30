'Can be called or assigned to a macro
public Sub OpenNoCode()
    Dim wbWithCode As String
    wbWithCode = InputBox("Path of workbook to open with code to suppress.", "Which Workbook?", CurDir & "\...xlsm")
    If wbWithCode <> vbNullString Then
        'stop Workbook_Open events from running in the workbook
        Application.EnableEvents = False
            Workbooks.Open wbWithCode
        Application.EnableEvents = True
    End With
End Sub