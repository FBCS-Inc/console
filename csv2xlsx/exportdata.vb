
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Public Class ExportData
	Private dt As DataTable
	Private filename As String
	Private appXL As Excel.Application = Nothing
	Private wbXl As Excel.Workbook = Nothing
	Private shXL As Excel.Worksheet = Nothing
	Private clXL As Excel.Range = Nothing
	Private seeit As Boolean
	Private ReadOnly spath As String
	Private startrow As Integer
	Private c As Integer
	Private x As Integer
	Private fmt As String
	Private asheet As String
	Private rowslimit As Integer
	Dim row As Integer
	Private ReadOnly col() As String = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	Public Sub New()
		Initialize()

	End Sub
	Public Sub Initialize()
		seeit = False
		startrow = 2
		c = 0
		x = 0
		row = startrow
		If rowslimit = 0 Then
			rowslimit = 1000000
		End If

	End Sub
	Public Property ActiveSheet() As String
		Get
			Return asheet
		End Get
		Set(value As String)
			asheet = value
		End Set
	End Property
	Public Property RowsPerSheet() As Integer
		Get
			Return rowslimit
		End Get
		Set(value As Integer)
			rowslimit = value
		End Set
	End Property
	Public Property ExcelFile() As String
		Get
			Return filename
		End Get
		Set(value As String)
			filename = value
		End Set
	End Property
	Public Property Data() As DataTable
		Get
			Return dt
		End Get
		Set(value As DataTable)
			dt = value
		End Set
	End Property
	Public Property Visable() As Boolean
		Get
			Return seeit
		End Get
		Set(value As Boolean)
			seeit = value
		End Set
	End Property
	Public Property FirstRow() As Integer
		Get
			Return startrow
		End Get
		Set(value As Integer)
			startrow = value
		End Set
	End Property

	Public WriteOnly Property SheetFormating As String

		Set(value As String)
			fmt = value
		End Set

	End Property
	' Start Excel and get Application object.
	Public Sub Export()
		If ExcelFile.EndsWith(".csv") Then
			Call ExportCSV()
		ElseIf ExcelFile.EndsWith(".xlsx") Then
			If File.Exists(ExcelFile) Then
				Call ExportExistingExcel()
			Else
				Call ExportNewExcel()
			End If
		Else
			MsgBox("Only XLSX and CSV can be exported at this time", 0, "Unable to export file")
		End If

	End Sub
	Private Sub ExportExistingExcel()

		Dim row As Integer = startrow
		Try
			appXL = CreateObject("Excel.Application")
			wbXl = appXL.Workbooks.Open(filename)
			shXL = wbXl.Worksheets(SheetName())

			appXL.DisplayAlerts = False
			'write headers to sheet
			c = 0
			For Each col As DataColumn In dt.Columns
				c += 1
				shXL.Cells(1, c) = col.ColumnName
			Next
			FormatSheet(shXL)
			'write data to sheet
			For Each dr As DataRow In dt.Rows
				For x = 1 To c
					If IsDBNull(dr(x - 1)) Then
						shXL.Cells(row, x) = ""
					Else
						shXL.Cells(row, x) = dr(x - 1)
					End If
				Next
				row += 1
				If row > rowslimit Then NextSheet()
			Next
			shXL.Columns.AutoFit()
			wbXl.SaveAs(filename, 51)
		Catch ex As Exception
			MsgBox("Can't Open excel " & filename & vbCrLf & ex.Message)
		Finally
			Marshal.ReleaseComObject(wbXl)
			wbXl = Nothing
			appXL.Quit()
			appXL.DisplayAlerts = True
			Marshal.ReleaseComObject(appXL)
			appXL = Nothing
		End Try

	End Sub
	Private Sub ExportNewExcel()
		Try
			'excel = New Excel.Application '.ApplicationClass
			appXL = CreateObject("Excel.Application")
			wbXl = appXL.Workbooks.Add()
			shXL = wbXl.Worksheets(SheetName())
			FormatSheet(shXL)
			appXL.DisplayAlerts = False
			'write headers to sheet
			c = 0
			For Each col As DataColumn In dt.Columns
				c += 1
				shXL.Cells(1, c) = col.ColumnName
			Next
			'write data to sheet
			For Each dr As DataRow In dt.Rows
				For x = 1 To c
					If IsDBNull(dr(x - 1)) Then
						shXL.Cells(row, x) = ""
					Else
						shXL.Cells(row, x) = dr(x - 1)
					End If
				Next
				row += 1
				If row > rowslimit Then NextSheet()
			Next
			shXL.Columns.AutoFit()
			wbXl.SaveAs(filename, 51)
			Marshal.ReleaseComObject(wbXl)
			wbXl = Nothing
		Catch ex As Exception
			MsgBox("Can't open excel " & filename & vbCrLf & ex.Message)
			Marshal.ReleaseComObject(appXL)
			appXL = Nothing
			wbXl = Nothing
			shXL = Nothing
		Finally
			appXL.Quit()
			appXL.DisplayAlerts = True
			Marshal.ReleaseComObject(appXL)
			appXL = Nothing
		End Try

	End Sub
	Private Sub NextSheet()
		'asheet += 1
		shXL = wbXl.Worksheets(SheetName())
		FormatSheet(shXL)
		'write headers to new sheet
		c = 0
		For Each col As DataColumn In dt.Columns
			c += 1
			shXL.Cells(1, c) = col.ColumnName
		Next
		row = startrow
	End Sub
	Private Sub ExportCSV()

		Dim sw As New StreamWriter(filename, False)
		Try
			'.csv file header
			c = 0
			For Each col As DataColumn In dt.Columns
				c += 1

				sw.Write(NoComma(col.ColumnName))
				If c < dt.Columns.Count Then
					sw.Write(",")
				Else
					sw.WriteLine("")
				End If
			Next

			For Each dr As DataRow In dt.Rows
				'Write CSV file
				For x = 0 To dt.Columns.Count - 1
					If IsDBNull(dr(x)) Then
						sw.Write("")
					Else
						sw.Write(NoComma(dr(x)))
					End If

					If x < dt.Columns.Count - 1 Then
						sw.Write(",")
					Else
						sw.WriteLine("")
					End If
				Next x

			Next dr

		Catch ex As Exception
			MsgBox("Can't open file " & filename & vbCrLf & ex.Message)
		Finally
			sw.Flush()
			sw.Close()
		End Try

	End Sub
	Private Function NoComma(txt As String) As String
		NoComma = txt.Replace(",", "")
	End Function
	Private Sub FormatSheet(ByRef shxl As Excel.Worksheet)
		shxl.Columns.AutoFit()
		If fmt = "" Then SetTextFMT()
		Dim fcode() As Char
		Dim xc As Integer = 0
		shxl.Select()

		For Each cf As String In fmt
			fcode = cf.ToCharArray
			For Each fc As Char In fcode
				Try
					clXL = shxl.Range(col(xc) & ":" & col(xc))
					clXL.Select()
					clXL.HorizontalAlignment = Excel.Constants.xlCenter
					clXL.VerticalAlignment = Excel.Constants.xlCenter
					Select Case fc
						Case "T"
							clXL.NumberFormat = "@"
						Case "C"
							clXL.NumberFormat = "$#,##0.00;[Red]$#,##0.00"
						Case "D"
							clXL.NumberFormat = "MM/dd/yyyy hh:mm:ss"
						Case "d"
							clXL.NumberFormat = "MM/dd/yyyy"
						Case "R"
							clXL.NumberFormat = "yyyy-MM-dd hh:mm:ss"
						Case "r"
							clXL.NumberFormat = "yyyy-MM-dd"
						Case "p"
							clXL.NumberFormat = "0.00""%"""
						Case "P"
							clXL.NumberFormat = "0.00%"
						Case Else
							clXL.NumberFormat = "@"
					End Select
					xc += 1
					clXL = Nothing
				Catch ex As Exception
					MsgBox("Can't format excel range " & col(xc) & ":" & col(xc) & " " & filename & vbCrLf & ex.Message)

				End Try


			Next
		Next
	End Sub
	Private Sub SetTextFMT()
		For c As Integer = 0 To dt.Columns.Count
			fmt += "T"
		Next
	End Sub
	Private Function SheetName() As String
		Dim sn As String, SheetNameTest As String
		sn = Trim(asheet.ToString)
		On Error Resume Next
		SheetNameTest = wbXl.Worksheets(sn).Name
		If Err.Number <> 0 Then
			Err.Clear()
			wbXl.Worksheets.Add.Name = sn
		End If
		Return sn

	End Function

End Class


