Sub Main
	Call DuplicateKeyExclusion()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Análisis: Excluir clave duplicada
Function DuplicateKeyExclusion
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.DupKeyExclusion
	task.IncludeAllFields
	task.AddKey "NUM_FACT", "A"
	task.DifferentField = "FECHA_FACT"
	task.Criteria = " COD_PROD  = ""05"""
	dbName = "Duplicado_02.IMD"
	task.CreateVirtualDatabase = False
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function