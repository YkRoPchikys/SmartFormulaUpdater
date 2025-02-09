Attribute VB_Name = "SmartFormulaUpdater"
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' ѕроверка, что изменени€ произошли в листе с формулами
    If Not Intersect(Target, Sh.UsedRange) Is Nothing Then
        Application.Calculate ' ѕересчитывает все формулы в книге
    End If
End Sub
