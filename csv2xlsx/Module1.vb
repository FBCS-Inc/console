Imports System.IO

Module Module1
	'sArg(0) is the input csv file including path
	'sArg(1) is the first row of data on the output excel after the header
	'sArg(2) is a text string for colum formatting 
	'T=text 
	'C=currenct 
	'D=datetime 
	'd=date
	'R=date reversed DateTime Year Month Day And time
	'r= Reversed date Year Month Day\
	'P=Percent it will multiply the value by 100 abnd add the % sign
	'p=Pwecent it will not change the column value but will add % to the column value
	'sArg(3) is the active sheet name
	'call this program once for every sheet you want in the excel output. 
	Sub Main(ByVal sArg() As String)
		'Dim lines = IO.File.ReadAllLines("C:\Users\Edward.Hall\Desktop\testfile.csv")
		Dim lines = IO.File.ReadAllLines(sArg(0))
		Dim tbl = New DataTable

		Dim cname As String() = lines.First.Split(","c)
		Dim colCount = cname.Length
		For i As Int32 = 0 To colCount - 1
			tbl.Columns.Add(New DataColumn(cname(i), GetType(String)))
		Next
		For Each line In lines
			If line <> lines.First Then
				Dim objFields = From field In line.Split(","c)
								Select CType(field, Object)
				Dim newRow = tbl.Rows.Add()
				newRow.ItemArray = objFields.ToArray()
			End If
		Next
		Dim xl = New ExportData
		With xl
			.Data = tbl
			.ExcelFile = Path.ChangeExtension(sArg(0), ".xlsx")
			'.ExcelFile = Path.ChangeExtension("C:\Users\Edward.Hall\Desktop\testfile.csv", ".xlsx")
			.FirstRow = sArg(1)
			.SheetFormating = sArg(2)
			.ActiveSheet = sArg(3)
		End With
		xl.Export()
	End Sub

End Module
