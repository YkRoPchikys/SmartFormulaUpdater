Автоматическое обновление формул в Excel (VBA макрос)
Описание
Этот макрос предназначен для автоматического обновления всех формул в рабочем листе при изменении данных. Он использует событие Workbook_SheetChange, чтобы отслеживать изменения в данных на любом из листов книги и автоматически пересчитывает все формулы с помощью команды Application.Calculate.

Макрос полезен, если вы часто работаете с данными в Excel и хотите, чтобы формулы всегда обновлялись без необходимости вручную их пересчитывать.

Как использовать
Откройте вашу книгу Excel.
Нажмите Alt + F11, чтобы открыть редактор VBA.
В левой части окна найдите ThisWorkbook и дважды кликните на него.
Вставьте код макроса в окно ThisWorkbook.
Сохраните книгу и убедитесь, что макросы включены.
Теперь всякий раз, когда вы вносите изменения в данные на любом листе, все формулы в книге будут автоматически пересчитаны.

Важное замечание
Этот макрос пересчитывает все формулы в книге. Если вы хотите ограничить обновление только на определенных листах или в определенных диапазонах, вы можете изменить код в соответствии с вашими потребностями.

Примечания
Поддерживается версия Excel с поддержкой VBA.
Если у вас возникнут вопросы, не стесняйтесь обращаться через Issues на GitHub.
Automatic Formula Update in Excel (VBA Macro)
Description
This macro is designed to automatically update all formulas in a worksheet when data is changed. It uses the Workbook_SheetChange event to track changes in the data on any sheet of the workbook and automatically recalculates all formulas using the Application.Calculate command.

This macro is useful if you frequently work with data in Excel and want your formulas to always be updated without the need to manually recalculate them.

How to Use
Open your Excel workbook.
Press Alt + F11 to open the VBA editor.
On the left, find ThisWorkbook and double-click on it.
Paste the macro code into the ThisWorkbook window.
Save the workbook and make sure macros are enabled.
Now, every time you make a change to the data on any sheet, all formulas in the workbook will be automatically recalculated.

Important Note
This macro recalculates all formulas in the workbook. If you want to limit the update to specific sheets or ranges, you can modify the code according to your needs.

Notes
Supported on Excel versions that support VBA.
If you have any questions, feel free to reach out via Issues on GitHub.