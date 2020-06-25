Sub Main
	Call DuplicateKeyDetection()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Análisis: Detectar clave duplicada
Function DuplicateKeyDetection
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.DupKeyDetection
	task.IncludeAllFields
	task.AddKey "NUM_FACT", "A"
	task.OutputDuplicates = TRUE
	dbName = "Duplicado_01.IMD"
	task.CreateVirtualDatabase = False
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function