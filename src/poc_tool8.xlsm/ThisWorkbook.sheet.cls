Private Sub Workbook_Open()

    With Worksheets(SheetNameSimple)
        .cbWindow.AddItem "�C��"
        .cbWindow.AddItem "�C3��"
        .cbWindow.AddItem "�C5��"
        .cbWindow.AddItem "�C��"
        
        .cbOccurrence.AddItem "�ܤ�1��"
        .cbOccurrence.AddItem "�ܤ�3��"
        .cbOccurrence.AddItem "�ܤ�5��"
        
        .cbOpponent.AddItem "�Ҧ��H"
        .cbOpponent.AddItem "���w�b��"
        '.cbOpponent.AddItem "���@�b��" ' TODO:
        
    End With
    
    menu.createMenu

End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Sh.name = SheetNameSimple Then
        Sh.Range("O1").Value = MainAccName
        Sh.Range("O2").Value = MainAccId
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    menu.deleteMenu
End Sub