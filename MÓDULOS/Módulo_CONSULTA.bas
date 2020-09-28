Attribute VB_Name = "Módulo_CONSULTA"
Option Explicit
Global BD As New ADODB.Connection

Sub ABRIRCONEXAO()

Dim CS As String
Dim ARQ As String
On Error Resume Next

ARQ = ThisWorkbook.Path & "\" & "BD SISTEMA DE CADASTRO.accdb;"

CS = "Provider=Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & ARQ _
& "Persist Security Info=False;"

BD.Close
BD.Open CS

End Sub
