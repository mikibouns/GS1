Attribute VB_Name = "Module1"
Sub Protect()
' Protect Macros
    On Error GoTo ErrorHandler
    Sheets("TDSheet").Select
    ActiveSheet.Protect Password:="278278", UserInterfaceOnly:=True
    Sheets("TDSheet").Select
    ActiveWindow.SelectedSheets.Visible = False
    ActiveWorkbook.Protect Password:="278278", Structure:=True, Windows:=False
    GoTo Ends:
ErrorHandler:
    MsgBox (" нига и лист уже защищены!")
Ends:
End Sub
Sub Unprotect()
' Unprotect Macros
    ActiveWorkbook.Unprotect Password:="278278"
    Sheets("TDSheet").Visible = True
    Sheets("TDSheet").Select
    ActiveSheet.Unprotect Password:="278278"
End Sub
