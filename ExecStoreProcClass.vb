imports System.IO
imports System.IO.Compression
Imports ADODB
public class StoreProcedure
	'-- Public varian
	Private Conn As ADODB.Connection
	Private sKey As String

	Public Function Encrypt(Byval s As String) As String
		Dim rijndaelEnhanced As RijndaelEnhanced
		rijndaelEnhanced = New RijndaelEnhanced(sKey, "@1B2c3D4e5F6g7H8")
		Return rijndaelEnhanced.Encrypt (s)
	End Function
	
	Public Function Decrypt(Byval s As String) As String
		Dim rijndaelEnhanced As RijndaelEnhanced
		rijndaelEnhanced = New RijndaelEnhanced(sKey, "@1B2c3D4e5F6g7H8")
		Return rijndaelEnhanced.Decrypt (s)
	End Function
	
	'-- b. DB
	Private Sub OpenConnDB()
		Conn = New ADODB.Connection
		Conn.Open(Decrypt(System.Configuration.ConfigurationManager.appSettings("ConnectionString").toString()))
	End Sub
	Private Sub CloseConnDB()
		Conn.Close
		Conn = nothing
	End Sub
	
	'-- struct
	Public Sub New(Byval strKey As String)
		sKey = strKey
	End Sub
	
	'-- Store dynamic
	Public Function ExecuteStore (Byval StoreName As String, Byval StrParameter As String, ByRef StrParameterOutput As String) As String
		Dim cmd As ADODB.Command, rs As ADODB.Recordset, Param (100) As ADODB.Parameter,
			aParameter1 () As String, aParameter () As String, ParamCount As Integer, i As Integer, j As Integer,
			aParamProperties (100,5) As String, ReturnString As String
		OpenConnDB
		aParameter = Split(StrParameter, "$")
		ParamCount = UBound(aParameter)
		For i = 0 To ParamCount
			aParameter1 = Split(aParameter (i), ";")
			For j = 0 To UBound (aParameter1)
				aParamProperties (i,j) = aParameter1(j)
			Next
			Param (i) = New ADODB.Parameter
		Next
		cmd = New ADODB.Command
		rs = New ADODB.Recordset
		cmd.ActiveConnection = Conn
		cmd.CommandType = 4
		cmd.CommandText = StoreName
		For i = 0 To ParamCount
			Param (i) = cmd.CreateParameter(aParamProperties (i,0), aParamProperties (i,1), aParamProperties (i,2), aParamProperties (i,3), aParamProperties (i,4))
			cmd.Parameters.Append (Param (i))
		Next
		StrParameterOutput = ""
		Try
			Dim kt As Boolean, sData As String
			rs = cmd.Execute
			sData = ""
			ReturnString = "{""ExecuteResponse"":""1"",""ExecuteMessage"":""Thành công"",""Data"":[DATASTRINGJSON]}"
			kt = (rs.State = 1)
			If (kt) Then 
				While (Not rs.EOF) 
					SData &= ",{""" & RS(0).Name & """:""" & RS(0).Value & """"
					For i = 1 To RS.Fields.Count - 1
						SData &= ",""" & RS(i).Name & """:""" & RS(i).Value & """"
					Next
					SData &= "}"					
					rs.MoveNext
				End While
				rs.Close
			End If
			If (Len(SData)>1) Then SData = RIGHT(SData, LEN(SData)-1)
			ReturnString = REPLACE(ReturnString, "DATASTRINGJSON", SData)
			For i = 0 To ParamCount
				If (aParamProperties (i,2) = 2 OR aParamProperties (i,2) = 3) Then
					StrParameterOutput &= ",""" & aParamProperties (i,0) & """:""" & Param (i).Value & """"
				End If
			Next
		Catch Ex As Exception
			ReturnString = "{""ExecuteResponse"":""-99"",""ExecuteMessage"":""" & Ex.ToString & """""Data"":[]}"
			For i = 0 To ParamCount
				If (aParamProperties (i,2) = 2 OR aParamProperties (i,2) = 3) Then
					StrParameterOutput &= ",""" & aParamProperties (i,0) & """:""-99." & Ex.ToString & """"
				End If
			Next
		End Try
		If (Len(StrParameterOutput)>1) Then StrParameterOutput = "{" & RIGHT(StrParameterOutput, LEN(StrParameterOutput)-1) & "}"
		CloseConnDB
		Return ReturnString
	End Function
end class
