Attribute VB_Name = "Module3"
Option Explicit

Private Sub Import_Hold()

Dim connection As New ADODB.connection

connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\tkinder\Desktop\Holds\query.iqy.xlsx; "
Extended Properties = "Excel 12.0 Xml;HDR=YES"

Dim query As String
query = "Select * From [Simple$]"

Dim rs As New ADODB.Recordset

rs.Open query, connection

connection.Close


End Sub

― Z πJ [ °· \ ΠΈ ]  ^ ° _ ¬ ` °E a π‘ b PΊ