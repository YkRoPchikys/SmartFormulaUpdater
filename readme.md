# SmartFormulaUpdater

## Описание

**SmartFormulaUpdater** — это макрос для автоматического обновления формул в рабочих листах Excel при изменении данных. Он помогает пользователям поддерживать актуальные вычисления без необходимости вручную пересчитывать формулы.

Этот макрос отслеживает изменения данных в любом из листов книги и автоматически обновляет все формулы, используя команду пересчета `Application.Calculate`. Это решение идеально подходит для больших рабочих книг, где обновление формул вручную может занять много времени.

## Установка

1. Откройте Excel и нажмите `Alt + F11`, чтобы открыть редактор VBA.
2. В панели слева найдите **ThisWorkbook** и дважды щелкните на него.
3. Вставьте следующий код в окно редактора:

```vba
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Проверка, что изменения произошли в листе с формулами
    If Not Intersect(Target, Sh.UsedRange) Is Nothing Then
        Application.Calculate ' Пересчитывает все формулы в книге
    End If
End Sub
Сохраните изменения и закройте редактор VBA.
Как использовать
Каждый раз, когда данные на любом из листов изменяются, макрос автоматически пересчитывает все формулы в книге.
Вам не нужно вручную запускать пересчет формул — это происходит автоматически при каждом изменении данных.
Особенности
Автоматическое обновление формул: После каждого изменения данных в книге, формулы обновляются без участия пользователя.
Подходит для больших файлов: Макрос идеально работает с большими рабочими книгами, где множество формул зависят от данных в различных ячейках.
Простота в установке: Все, что нужно — это добавить код в ThisWorkbook, и макрос сразу начнёт работать.
Лицензия
Этот проект распространяется под лицензией MIT. Подробнее см. файл LICENSE.

SmartFormulaUpdater
Description
SmartFormulaUpdater is a macro designed to automatically update formulas in Excel worksheets when data changes. It helps users maintain accurate calculations without the need to manually recalculate formulas.

This macro tracks data changes in any sheet of the workbook and automatically updates all formulas by using the Application.Calculate command. This solution is perfect for large workbooks where manually updating formulas could be time-consuming.

Installation
Open Excel and press Alt + F11 to open the VBA editor.
In the left panel, find ThisWorkbook and double-click it.
Paste the following code into the editor window:
vba
Копировать
Редактировать
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Check if the changes happened in the sheet with formulas
    If Not Intersect(Target, Sh.UsedRange) Is Nothing Then
        Application.Calculate ' Recalculates all formulas in the workbook
    End If
End Sub
Save the changes and close the VBA editor.
How to Use
Every time data changes on any sheet, the macro automatically recalculates all formulas in the workbook.
You don't need to manually trigger the formula recalculation — it happens automatically when the data changes.
Features
Automatic formula update: After each data change in the workbook, formulas are updated without user interaction.
Works well with large files: The macro performs excellently with large workbooks where many formulas depend on data from various cells.
Easy installation: All you need to do is add the code to ThisWorkbook, and the macro will start working right away.
License
This project is licensed under the MIT License. See the LICENSE file for more information.