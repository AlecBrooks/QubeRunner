
set WshShell = createobject("Wscript.shell")
set objexcel = createobject("Excel.application")
objexcel.visible = false
set objWorkbook = objexcel.workbooks.Open(WshShell.CurrentDirectory & "\QubeRunner0_2.xlsm")
