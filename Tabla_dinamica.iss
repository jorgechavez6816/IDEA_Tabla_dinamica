Sub Main
	Call PivotTable()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Análisis: Tabla dinámica
Function PivotTable
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.PivotTable
	task.ResultName = "Tabla_dinamica_011"
	task.AddRowField "NUM_VENDEDOR"
	task.AddColumnField "COD_PROD"
	task.AddDataField "TOTAL", "Suma: TOTAL", 1
	task.ExportToIDEA False
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function