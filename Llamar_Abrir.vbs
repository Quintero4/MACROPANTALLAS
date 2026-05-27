Dim ExcelApp, Workbook
Set ExcelApp = CreateObject("Excel.Application")
' MODIFICA LA RUTA DE ABAJO POR LA DE TU EXCEL REAL:
Set Workbook = ExcelApp.Workbooks.Open("\\Spi60001.eadscasa.casa.corp\a400m san pablo\Lean_Manufacturing\4.- Innovation\3. Projects info\13. Visual Plant\Pantallas N2\MacroPantallasSTs.xlsm")
ExcelApp.Run "ThisWorkbook.AbrirVideoDinamico"
Workbook.Close False
ExcelApp.Quit