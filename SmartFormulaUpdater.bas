Attribute VB_Name = "SmartFormulaUpdater"
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' ��������, ��� ��������� ��������� � ����� � ���������
    If Not Intersect(Target, Sh.UsedRange) Is Nothing Then
        Application.Calculate ' ������������� ��� ������� � �����
    End If
End Sub
