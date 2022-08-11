Dim filePath As String
Dim file As String
Sub Main
	file = "BOBI FRS NAT Dolayli Konsolide"
	Call ImportingExcel()
	Call ExtractFrom()
	Call FieldCreation()
	Call appendDb()
	Call createDb()
	Call appendDbfinal()
End Sub

Function ImportingExcel()
	' Create the task.
	Set task = Client.GetImportTask("ImportExcel")
	dbName =  "C:\Users\Arda\Desktop\Med-Idea\Red\Red_Assignment.xlsx"
	task.FileToImport = dbName
	task.SheetToImport = file
	task.OutputFilePrefix = "Arda"
	task.FirstRowIsFieldName = "True"
	task.EmptyNumericFieldAsZero = "True"
	' Perform the task.
	task.PerformTask
	' Obtain the output file name.
	dbName = task.OutputFilePath(file)
	' Clear the memory.
	Set task = Nothing
	' Open the result.
	Client.OpenDatabase(dbName)
End Function

Function ExtractFrom()
	' Open the database.
	Set db = Client.OpenDatabase("Arda-BOBI FRS NAT Dolayli Konsolide.IMD")
	
	' Create the task.
	
	Set task = db.Extraction
	
	' Configure the task.
	
	task.addFieldToInc "COL3"
	
	dbName = "Extraction.IMD"
	
	task.AddExtraction dbName, "", "@IsBlank(COL3)=0"
	
	' Perform the task.
	
	task.PerformTask 1, db.Count
	
	' Clear the memory.
	
	Set task = Nothing
	
	Set db = Nothing

	' Open the result.
	
	Client.OpenDatabase (dbName)
	' Open the database.
	Set db = Client.OpenDatabase("Arda-BOBI FRS NAT Dolayli Konsolide.IMD")
	
	' Create the task.
	
	Set task = db.Extraction
	
	' Configure the task.
	
	task.addFieldToInc "BOBI_FRS_NAKIT_AKIÞ_TABLOSU_DOLAYLI_YÖNTEM_"
	
	dbName = "Extraction2.IMD"
	
	task.AddExtraction dbName, "", "@IsBlank(BOBI_FRS_NAKIT_AKIÞ_TABLOSU_DOLAYLI_YÖNTEM_)=0"
	
	' Perform the task.
	
	task.PerformTask 1, db.Count
	
	' Clear the memory.
	
	Set task = Nothing
	
	Set db = Nothing

	' Open the result.
	
	Client.OpenDatabase (dbName)
	
	' Open the database.
	Set db = Client.OpenDatabase("Arda-BOBI FRS NAT Dolayli Konsolide.IMD")
	
	' Create the task.
	
	Set task = db.Extraction
	
	' Configure the task.
	
	task.addFieldToInc "temp2"
	
	dbName = "Extraction3.IMD"
	
	task.AddExtraction dbName, "", "@IsBlank(temp2)=0"
	
	' Perform the task.
	
	task.PerformTask 1, db.Count
	
	' Clear the memory.
	
	Set task = Nothing
	
	Set db = Nothing
	
	' Open the result.
	
	Client.OpenDatabase (dbName)
End Function


Function FieldCreation
	'
	Set db = Client.OpenDatabase("Extraction.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ACIKLAMALAR"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 88
	task.ReplaceField "temp1", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
	Set db = Client.OpenDatabase("Extraction2.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ACIKLAMALAR"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 88
	task.ReplaceField "BOBI_FRS_NAKIT_AKIS_TABLOSU_DOLAYLI_YÖNTEM_", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
	Set db = Client.OpenDatabase("Extraction3.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ACIKLAMALAR"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 88
	task.ReplaceField "temp2", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

Function appendDb

	Set db = Client.OpenDatabase("Extraction.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "Extraction2.IMD"
	task.AddDatabase "Extraction3.IMD"
	dbName = "Append.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

Function createDb

	Dim NewTable As Table
	Set NewTable = Client.NewTableDef
	
	' Define a field for the table.
	Dim AddedField As Field
	Set AddedField = NewTable.NewField
	AddedField.Name = "ACIKLAMALAR"
	AddedField.Type = WI_CHAR_FIELD
	AddedField.Length = 200
	
	' Add the field to the table.
	NewTable.AppendField AddedField
	
	' Perform the same steps for a second field.
	Set AddedField = NewTable.NewField
	AddedField.Name = "CARI_DONEM"
	AddedField.Type = WI_NUM_FIELD
	AddedField.Decimals = 1
	NewTable.AppendField AddedField
	
	
	Set AddedField = NewTable.NewField
	AddedField.Name = "ONCEKI_DONEM"
	AddedField.Type = WI_NUM_FIELD
	AddedField.Decimals = 1
	NewTable.AppendField AddedField
	
	' Change the table settings to allow writing.
	NewTable.Protect = False
	
	' Create the database.
	Dim db As Database
	Set db = Client.NewDatabase("Arda_Ciftci_test.IMD", "", NewTable)
	
	' Obtain the recordset.
	Dim rs As RecordSet
	Set rs = db.RecordSet
	
	' Obtain a new record.
	Dim rec As Record
	Set rec = rs.NewRecord
	
	' Use the field name method to add data.
	rec.SetCharValue "ACIKLAMALAR"," "
	rec.SetCharValue "CARI_DONEM", 0
	rec.SetCharValue "ONCEKI_DONEM", 0

	rs.AppendRecord rec
	
	' Protect the table before you commit it.
	NewTable.Protect = True
	
	' Commit the database.
	db.CommitDatabase
	' Open the database.
	Client.OpenDatabase "Arda_Ciftci_test.IMD"
	' Clear the memory.
	
	Set db = Nothing
	Set AddedField = Nothing
	Set NewTable = Nothing
End Function

Function appendDbfinal
	Set db = Client.OpenDatabase("Arda_Ciftci_test.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "Append2.IMD"
	task.Criteria = "@IsBlank(ACIKLAMALAR)=0"
	dbName = "Arda_Ciftci_test.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function