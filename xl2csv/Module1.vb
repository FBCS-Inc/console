
Imports System.IO
Imports Microsoft
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic.FileIO

Module Module1

	Sub Main(ByVal sArgs() As String)
		Dim FileName = "C:\Users\Edward.Hall\Desktop\ConnsDemoApplication.xlsx"
		'Dim row As Integer = 1
		If sArgs.Length > 0 Then
			'Dim FileName = sArgs(0)
			Dim row = 1 'CInt(sArgs(1))
			Dim excel As New Excel.Application With {
				.DisplayAlerts = False
			}
			Dim workbook As Excel.Workbook = excel.Workbooks.Open(FileName)
			Dim sheet As Excel.Worksheet = workbook.Sheets("sheet1")
			If row = 99 Then
				FileName = Path.ChangeExtension(FileName, ".out")
			Else
				FileName = Path.ChangeExtension(FileName, ".csv")
			End If

			sheet.SaveAs(FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows)

			workbook.Close()
			workbook = Nothing

			excel.Quit()
			excel = Nothing
			If row <> 99 Then
				Startrow(row, FileName)
			End If
		End If
	End Sub
	Private Sub Startrow(ByVal r As Integer, ByVal strFilePath As String)

		Dim filename As String = Path.ChangeExtension(strFilePath, ".out")
		Dim tfp As New TextFieldParser(strFilePath)
		tfp.Delimiters = New String() {","}
		tfp.TextFieldType = FieldType.Delimited
		If r > 0 Then
			For x = 0 To r - 1
				tfp.ReadLine() ' skip header
			Next
		End If
		While tfp.EndOfData = False
			Dim fields = tfp.ReadFields()
			Dim line As String = Nothing
			For Each Dta In fields
				If Dta.Contains("""") Then
					line += Dta.ToString
				Else
					If line Is Nothing Then
						line += """" + Dta.ToString + """"
					Else
						line += ",""" + Dta.ToString + """"
					End If
				End If
			Next
			Using writer As New StreamWriter(filename, True)
				writer.WriteLine(line)
			End Using
		End While
		tfp.Close()
		My.Computer.FileSystem.DeleteFile(Path.ChangeExtension(strFilePath, ".csv"))
	End Sub
End Module
