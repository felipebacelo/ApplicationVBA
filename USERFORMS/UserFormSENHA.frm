VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSENHA 
   Caption         =   "GERENCIADOR DE SENHAS"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8565
   OleObjectBlob   =   "UserFormSENHA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSENHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBoxSENHA_Click()

If CheckBoxSENHA = True Then
    TextBoxSENHAATUAL.PasswordChar = ""
    TextBoxNOVASENHA.PasswordChar = ""
    TextBoxCONFSENHA.PasswordChar = ""
Else
    TextBoxSENHAATUAL.PasswordChar = "*"
    TextBoxNOVASENHA.PasswordChar = "*"
    TextBoxCONFSENHA.PasswordChar = "*"
End If

End Sub

Private Sub CommandButtonCANCELSENHA_Click()

Unload Me

End Sub

Private Sub CommandButtonOKSENHA_Click()

Dim I As Integer
Dim USUARIOATUAL As String
Dim SENHA As String, SENHAATUAL As String
Dim SQL As String
Dim RS As New ADODB.Recordset


USUARIOATUAL = TextBoxUSUARIOS.Value
SENHA = TextBoxSENHAATUAL.Value

SQL = "SELECT SENHA FROM USUARIO WHERE USUÁRIO LIKE " & "'" & USUARIOATUAL & "'"
RS.Open SQL, BD

SENHAATUAL = RS.Fields(0).Value
RS.Close

Select Case ""
    Case Is = TextBoxSENHAATUAL.Value, TextBoxNOVASENHA.Value, TextBoxCONFSENHA.Value
    MsgBox "PREENCHA TODOS OS CAMPOS!", vbExclamation, "ATENÇÃO"
    Exit Sub
End Select

    If SENHAATUAL <> SENHA Then
        MsgBox "A SENHA ATUAL NÃO CONFERE!", vbCritical, "ATENÇÃO"
        TextBoxSENHAATUAL.SetFocus
        Exit Sub
    End If
    
    If Len(TextBoxNOVASENHA.Value) < 4 Then
        MsgBox "A NOVA SENHA DEVE TER NO MÍNIMO 4 CARACTERES!", vbExclamation, "ATENÇÃO"
        Exit Sub
    End If
    
    If TextBoxNOVASENHA.Value <> TextBoxCONFSENHA.Value Then
        MsgBox "AS NOVAS SENHAS NÃO CONFEREM", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    
SQL = "UPDATE USUARIO SET SENHA = " & "'" & TextBoxNOVASENHA.Value & "'"
SQL = SQL & " WHERE USUÁRIO LIKE " & "'" & USUARIOATUAL & "'"
RS.Open SQL, BD

MsgBox "SENHA ALTERADA COM SUCESSO!", vbInformation, "INFORMAÇÃO"
Unload Me
    
End Sub

Private Sub TextBoxCONFSENHA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub TextBoxNOVASENHA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub UserForm_Initialize()

TextBoxUSUARIOS.Locked = True
TextBoxUSUARIOS.Value = BDUSUARIOS.Range("USUARIOATUAL").Value

End Sub
