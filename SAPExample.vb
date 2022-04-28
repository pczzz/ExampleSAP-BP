Module ExampleSAP

	Sub Main()

		Dim session, mygrid As Object
		Dim sID, colName As String
		Dim rCount, cCount As Integer
		Dim r, c As Integer
		Dim table As DataTable = New DataTable()

		'=== Cast Element ID to string in order to recognize it in BP ===
		bSuccess = False
		sID = sSapId

		'=== Access SAP Scripting Engine ===
		session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
		mygrid = session.findById(sID)

		Try
			'=== Get Row and Column Count of SAP Table ===
			rCount = mygrid.RowCount
			cCount = mygrid.ColumnCount

			'=== Create Columns in DataTable according to SAP Table ===	
			For c = 0 To cCount - 1
				If bTechnicalHeaders Then
					colName = mygrid.columnorder(c)
				Else
					colName = mygrid.columnorder(c)
					colName = mygrid.getDisplayedColumnTitle(colName)
				End If
				table.Columns.Add(colName, System.Type.GetType("System.String"))
			Next c

			'=== Fill in the DataTable from SAP Table cell by cell ===	
			For r = 0 To rCount - 1
				mygrid.firstVisibleRow = r
				table.Rows.Add()

				For c = 0 To cCount - 1
					colName = mygrid.columnorder(c)
					table.Rows(r).Item(c) = mygrid.getcellvalue(r, colName)
				Next c
			Next r

			'=== Create Output Collection ===	
			Collection = table
			bSuccess = True

		Catch
			Throw New System.Exception("Robot was not able to transform SAP table into Collection.")
		End Try

	End Sub

End Module
