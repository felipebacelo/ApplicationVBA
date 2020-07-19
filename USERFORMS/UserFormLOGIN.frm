VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormLOGIN 
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5265
   OleObjectBlob   =   "UserFormLOGIN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormLOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBoxLOGIN_Click()

    If CheckBoxLOGIN.Value = True Then
        TextBoxLOGINSENHA.PasswordChar = ""
    Else
        If TextBoxLOGINSENHA.Value <> "SENHA" Then
            TextBoxLOGINSENHA.PasswordChar = "*"
        End If
    End If
   
End Sub

Private Sub CommandButtonLOGIN_Click()

Dim USUARIOLOGIN As String
Dim SENHALOGIN As String
Dim SQL As String
Dim RS As New ADODB.Recordset
On Error GoTo PULA:

USUARIOLOGIN = TextBoxLOGINUSUARIO.Value
SENHALOGIN = TextBoxLOGINSENHA.Value

SQL = "SELECT USUÁRIO,SENHA, NÍVEL FROM USUARIO "
SQL = SQL & " WHERE USUÁRIO LIKE " & "'" & USUARIOLOGIN & "'"

RS.Open SQL, BD, adOpenStatic

    If RS.RecordCount = 0 Then
        MsgBox "USUÁRIO INCORRETO!", vbCritical, "ATENÇÃO"
        RS.Close
        TextBoxLOGINUSUARIO.SetFocus
    Else
        If SENHALOGIN <> RS.Fields(1).Value Then
            MsgBox "SENHA INCORRETA!", vbCritical, "ATENÇÃO"
            RS.Close
            TextBoxLOGINSENHA.SetFocus
        End If
    End If

BDUSUARIOS.Range("USUARIOATUAL").Value = RS.Fields(0).Value
BDUSUARIOS.Range("NIVELATUAL").Value = RS.Fields(2).Value

BD.Close
Unload Me
UserFormSISTEMA.Show
PULA:
End Sub

Private Sub TextBoxLOGINSENHA_AfterUpdate()

    If TextBoxLOGINSENHA.Value = "" Then
        TextBoxLOGINSENHA.Value = "SENHA"
        TextBoxLOGINSENHA.PasswordChar = ""
    End If

End Sub

Private Sub TextBoxLOGINSENHA_Change()

    If CheckBoxLOGIN.Value = False Then
        TextBoxLOGINSENHA.PasswordChar = "*"
    End If

End Sub

Private Sub TextBoxLOGINSENHA_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If TextBoxLOGINSENHA.Value = "SENHA" Then
        TextBoxLOGINSENHA.Value = ""
    End If

End Sub

Private Sub TextBoxLOGINUSUARIO_AfterUpdate()

    If TextBoxLOGINUSUARIO.Value = "" Then
        TextBoxLOGINUSUARIO.Value = "USUÁRIO"
    End If

End Sub

Private Sub TextBoxLOGINUSUARIO_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If TextBoxLOGINUSUARIO.Value = "USUÁRIO" Then
        TextBoxLOGINUSUARIO.Value = ""
    End If

End Sub

Private Sub UserForm_Initialize()

Call ABRIRCONEXAO

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = 0 Then
    If Workbooks.Count = 1 Then
        Application.Visible = True
        Application.Quit
    Else
        Application.Visible = True
        ThisWorkbook.Close False
    End If
End If
    
End Sub
