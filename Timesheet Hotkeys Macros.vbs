Sub Add_dash()
'
' add dashes in front of the notes for neat organization before adding to an invoice
' Recommended keyboard Shortcut: Ctrl+d
'


Dim c As Range
For Each c In Selection
If c.Value <> "" Then c.Value = "- " & c.Value
Next


End Sub
Sub Send_To_Archive()
'
' Sends the currently highlighted row to a new sheet for archiving
' Recommended keyboard Shortcut: Ctrl+q
'
    Selection.Cut
    Sheets("Labor Archive").Select
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    Sheets("Labor and expense to bill").Select
End Sub
Sub Macro6()
'
' Delete the current row after it's been copied to the other sheet during archiving
'

'
    Selection.Delete Shift:=xlUp
End Sub
