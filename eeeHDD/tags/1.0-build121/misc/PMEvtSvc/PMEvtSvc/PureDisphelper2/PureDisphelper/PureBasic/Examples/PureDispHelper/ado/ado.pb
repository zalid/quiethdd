; example by Kiffi

EnableExplicit

Global Databasename.s = "ado.mdb"

XIncludeFile "adoconstants.pbi"

dhToggleExceptions(#True)

Procedure ADOX_Example_Create_New_Database_Append_Tables_And_Columns()

  If FileSize(Databasename) > 0
    If DeleteFile(Databasename) = 0
      ProcedureReturn
    EndIf
  EndIf

  Protected oCatalog.l
  Protected oTable.l
  Protected oColumn.l
  Protected ConnectionString.s

  ; Create the required ADOX-Object
  oCatalog = dhCreateObject("ADOX.Catalog")
  If oCatalog = 0
    MessageRequester("ADOX-Example", "Couldn't create ADOX.Catalog")
    ProcedureReturn
  EndIf

  ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Databasename + ";Jet OLEDB:Engine Type=5;"
  ; Engine Type=4: <= Access 97 Database (JET 3.5)
  ; Engine Type=5: >= Access 2000 Database (JET 4.0)
  ; Standardvalue: Engine Type=5

  If dhCallMethod(oCatalog, ".Create(%T)", @ConnectionString) <> 0
    MessageRequester("ADOX-Example", "Couldn't create Database")
    dhReleaseObject(oCatalog)
    ProcedureReturn
  EndIf

  ; Create a new table
  oTable = dhCreateObject("ADOX.Table")

  If oTable = 0
    MessageRequester("ADOX-Example", "Couldn't create table")
    dhReleaseObject(oCatalog)
    ProcedureReturn
  EndIf

  ; give the table a name
  dhPutValue(oTable, ".Name = %T", @"TestTable")

  ; Append table to catalog
  dhCallMethod(oCatalog, ".Tables.Append(%o)", oTable)

  ; Create an ID-Column with Autoincrement
  oColumn = dhCreateObject("ADOX.Column")
  dhPutValue(oColumn, ".ParentCatalog = %o", oCatalog)
  dhPutValue(oColumn, ".Name = %T", @"ID")
  dhPutValue(oColumn, ".Type = %d", #adInteger)
  dhPutValue(oColumn, ".Properties(%T) = %b", @"Autoincrement", #True)

  ; Append the new column to the table
  dhCallMethod(oTable, ".Columns.Append(%o)", oColumn)

  ; Create and append some more Columns
  dhCallMethod(oTable, ".Columns.Append(%T, %d)", @"IntegerField", #adInteger)
  dhCallMethod(oTable, ".Columns.Append(%T, %d)", @"TextField",    #adVarWChar)
  dhCallMethod(oTable, ".Columns.Append(%T, %d)", @"MemoField",    #adLongVarWChar)


  dhReleaseObject(oColumn)
  dhReleaseObject(oTable)
  dhReleaseObject(oCatalog)

  MessageRequester("ADOX-Example", "Ready. Database successfully created.")

EndProcedure

Procedure ADO_Example_Open_Existing_Table_And_Fill_In_Some_Data()

  Protected oCN.l ; Connection-Object

  Protected ConnectionString.s
  Protected SQL.s
  Protected Counter.l

  ; Connection string.
  ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Databasename + ";Jet OLEDB:Engine Type=5;"
  ; Engine Type=4: <= Access 97 Database (JET 3.5)
  ; Engine Type=5: >= Access 2000 Database (JET 4.0)
  ; Standardvalue: Engine Type=5

  ; Create the required ADO-Connection-Object
  oCN = dhCreateObject("ADODB.Connection")

  If oCN = 0
    MessageRequester("ADOX-Example", "Couldn't create ADODB.Connection")
    ProcedureReturn
  EndIf

  ; Open the connection.
  If Not dhCallMethod(oCN, ".Open(%T)", @ConnectionString)

    ; Insert some records.
    For Counter = 0 To 99
      SQL = "Insert Into TestTable (IntegerField, TextField, MemoField) Values (" + Str(Counter) + ", 'Text" + Str(Counter) + "', 'Memo" + Str(Counter) + "')"
      dhCallMethod(oCN, ".Execute(%T)", @SQL)
    Next

    ; Close the connection.
    dhCallMethod(oCN, ".Close")

    MessageRequester("ADOX-Example", "Ready. Database successfully filled with sample-data.")

  Else

    MessageRequester("ADOX-Example", "Couldn't open the Connection")

  EndIf

  dhReleaseObject(oCN)

EndProcedure

Procedure ADO_Example_Open_Existing_Table_And_Read_Data()

  Protected oRS .l ; Connection-Object
  Protected oCN.l ; Connection-Object

  Protected szResponse.l

  Protected ConnectionString.s
  Protected SQL.s
  Protected EOF.l

  ; Connection string.
  ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Databasename + ";Jet OLEDB:Engine Type=5;"
  ; Engine Type=4: <= Access 97 Database (JET 3.5)
  ; Engine Type=5: >= Access 2000 Database (JET 4.0)
  ; Standardvalue: Engine Type=5

  ; Create the required ADO-Connection-Object
  oCN = dhCreateObject("ADODB.Connection")
  If oCN = 0
    MessageRequester("ADO", "Couldn't create Connection-Object")
    ProcedureReturn
  EndIf

  ; Open the connection.
  If Not dhCallMethod(oCN, ".Open(%T)", @ConnectionString)

    ; Retrieve some records.
    SQL = "Select * from TestTable"

    ; Execute the query
    dhGetValue("%o", @oRS, oCN, ".Execute(%T)", @SQL)

    Repeat

      dhGetValue("%b", @EOF, oRS, ".EOF")

      If EOF
        Break
      EndIf

      dhGetValue("%T", @szResponse, oRS, ".Fields(%T).Value", @"ID")
      If szResponse : Debug "ID: " + PeekS(szResponse) : EndIf
      dhFreeString(szResponse) : szResponse = 0

      dhGetValue("%T", @szResponse, oRS, ".Fields(%T).Value", @"IntegerField")
      If szResponse : Debug "IntegerField: " + PeekS(szResponse) : EndIf
      dhFreeString(szResponse) : szResponse = 0

      dhGetValue("%T", @szResponse, oRS, ".Fields(%T).Value", @"TextField")
      If szResponse : Debug "TextField: " + PeekS(szResponse) : EndIf
      dhFreeString(szResponse) : szResponse = 0

      dhGetValue("%T", @szResponse, oRS, ".Fields(%T).Value", @"MemoField")
      If szResponse : Debug "MemoField: " + PeekS(szResponse) : EndIf
      dhFreeString(szResponse) : szResponse = 0

      Debug "-------"

      dhCallMethod(oRS, ".MoveNext")

    ForEver

    ; Close and release the recordset.
    dhCallMethod(oRS, ".Close")
    dhReleaseObject(oRS)

    ; Close and release the connection.
    dhCallMethod(oCN, ".Close")
    dhReleaseObject(oCN)

  EndIf

EndProcedure

ADOX_Example_Create_New_Database_Append_Tables_And_Columns()
ADO_Example_Open_Existing_Table_And_Fill_In_Some_Data()
ADO_Example_Open_Existing_Table_And_Read_Data()

