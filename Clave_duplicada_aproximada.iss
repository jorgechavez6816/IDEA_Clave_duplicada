Sub Main
	Call FuzzyDuplicate()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Análisis: Duplicación aproximada
Function FuzzyDuplicate
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.FuzzyDuplicate
	task.IncludeAllFields
	task.AddMatchField "NUM_FACT"
	task.Criteria = " COD_PROD  = ""05"""
	dbName = "Duplicado_03.IMD"
	task.OutputDBName = dbName
	task.CreateVirtualDatabase = False
	task.AllowRecordsInMultipleFuzzyGroups = True
	task.IncludeExactDuplicates = True
	task.MatchCase = False
	task.SimilarityDegreeThreshold = 0.8
	task.OutputType = WI_FD_OUTPUT_ALLRECORDS
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function