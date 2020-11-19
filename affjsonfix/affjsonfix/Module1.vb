Imports System.Text.RegularExpressions
Imports System.IO

Module Module1

	Sub Main(ByVal sargs() As String)
		Dim inputfile As String
		inputfile = "H:\New Business\AFF\2020\November\20201116\AccountUpdates_20201114_010219.json"
		Dim fileReader As String
		fileReader = My.Computer.FileSystem.ReadAllText(inputfile)
		Dim f As String

		f = Chr(34) + "originator" + Chr(34) + ":{" + Chr(34) + "servicer"
		'f = "originator" + ":" + " {" + """servicer""" + ":"
		Dim inx As Integer
		Dim text As String
		'inx = InStr(fileReader.ToLower, f.ToLower)
		'Dim text As String = fileReader.Substring(inx - 2, 64)
		For Each match As Match In Regex.Matches(fileReader.ToLower, f.ToLower)
			inx = InStr(fileReader.ToLower, f.ToLower)
			If inx > 0 Then
				text = fileReader.Substring(inx, 63)
			fileReader = fileReader.Replace(text, """originator""" + ":" + """creditor""")
			inx = 0
				text = ""
			End If
		Next

		Dim path As String = System.IO.Path.GetDirectoryName(inputfile)
		path += "\revision"

		If Not Directory.Exists(path) Then
			Directory.CreateDirectory(path)
		End If
		path += "\" + System.IO.Path.GetFileName(inputfile)
		Dim fileout As New IO.StreamWriter(path)
		fileout.Write(fileReader)
		fileout.Close()

	End Sub



End Module
