VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSISTEMA 
   Caption         =   "SISTEMA DE CADASTRO"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12885
   OleObjectBlob   =   "UserFormSISTEMA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSISTEMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButtonCADASTRAR_Click()

MultiPageGERAL.Value = 0
LBMODO.Caption = "NOVO CADASTRO"
CommandButtonLIMPAR_Click
CommandButtonLIMPARFILTRO_Click
LBID.Caption = CADASTROS.Range("ID").Value

End Sub

Private Sub CommandButtonCONSULTAR_Click()

MultiPageGERAL.Value = 1
CommandButtonLIMPARFILTRO_Click
CommandButtonFILTRAR_Click

End Sub

Private Sub CommandButtonDELETAR_Click()
Dim NIVEL As Integer

NIVEL = BDUSUARIOS.Range("NIVELATUAL").Value

    If NIVEL < 4 Then
        MsgBox "ESTE USUÁRIO NÃO POSSUI PERMISSÃO PARA CONFIGURAR USUÁRIOS!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

If ListBoxUSUARIOS.ListIndex = -1 Then Exit Sub

Dim RESPOSTA As VbMsgBoxResult

RESPOSTA = MsgBox("TEM CERTEZA DE QUE DESEJA DELETAR ESTE USUÁRIO?", vbYesNo + vbQuestion, "ATENÇÃO")

If RESPOSTA = vbYes Then Call DELETARUSUARIO(ListBoxUSUARIOS.Value)

CommandButtonUSUARIOS_Click

End Sub

Private Sub CommandButtonDELETAR1_Click()
Dim NIVEL As Integer
Dim RESPOSTAS As VbMsgBoxResult

NIVEL = BDUSUARIOS.Range("NIVELATUAL").Value

    If NIVEL < 3 Then
        MsgBox "ESTE USUÁRIO NÃO POSSUI PERMISSÃO PARA DELETAR CADASTROS!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

If UserFormSISTEMA.LBMODO.Caption = "NOVO CADASTRO" Then Exit Sub

RESPOSTA = MsgBox("TEM CERTEZA DE QUE DESEJA DELETAR ESTE CADASTRO?", vbYesNo + vbQuestion, "ATENÇÃO")

If RESPOSTA = vbYes Then Call DELETAR: CommandButtonCADASTRAR_Click

End Sub

Private Sub CommandButtonEDITARFILTRO_Click()

Dim ID As Long
Dim RS As New ADODB.Recordset
Dim SQL As String
Dim I As Integer
MultiPageGERAL.Value = 0

On Error Resume Next
ID = LISTAREGISTROS.Value

SQL = "SELECT * FROM CADASTROS "
SQL = SQL & " WHERE ID LIKE " & ID
RS.Open SQL, BD

    For I = 1 To 81
        CAMPO(I).Value = RS.Fields(I).Value
        CAMPO(I).BackColor = &H80000005
    Next

LBMODO.Caption = "EDITAR": LBID.Caption = ID

End Sub

Private Sub CommandButtonEDITARUSUARIO_Click()
If ListBoxUSUARIOS.ListIndex = -1 Then Exit Sub
Dim NIVEL As Integer

NIVEL = BDUSUARIOS.Range("NIVELATUAL").Value

    If NIVEL < 4 Then
        MsgBox "ESTE USUÁRIO NÃO POSSUI PERMISSÃO PARA CONFIGURAR USUÁRIOS!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

EDITARUSUARIO = True
UserFormUSUARIO.Show
CommandButtonUSUARIOS_Click

End Sub

Private Sub CommandButtonFILTRAR_Click()

Call FILTROAVANÇADO

End Sub

Private Sub CommandButtonLIMPAR_Click()

Call ATRIBUIRCAMPOS

For I = 1 To 81
    CAMPO(I).Value = ""
Next

End Sub

Private Sub CommandButtonLIMPARFILTRO_Click()

TextBoxFILTRONUCLEO.Value = ""
TextBoxFILTROQUADRA.Value = ""
TextBoxFILTROLOTE.Value = ""
TextBoxFILTRODOMICILIO.Value = ""
TextBoxFILTROENTREVISTADOR.Value = ""

End Sub

Private Sub CommandButtonNOVO_Click()
Dim NIVEL As Integer

NIVEL = BDUSUARIOS.Range("NIVELATUAL").Value

    If NIVEL < 4 Then
        MsgBox "ESTE USUÁRIO NÃO POSSUI PERMISSÃO PARA CONFIGURAR USUÁRIOS!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If

EDITARSUSUARIO = False
UserFormUSUARIO.Show
CommandButtonUSUARIOS_Click

End Sub

Private Sub CommandButtonSAIR_Click()

Application.Visible = True
ThisWorkbook.Save
Unload Me

End Sub

'CAMPOS OBRIGATÓRIOS'
Private Sub CommandButtonSALVAR_Click()

Dim I As Integer
Dim CAMPOSOB As Boolean
Dim NIVEL As Integer

NIVEL = BDUSUARIOS.Range("NIVELATUAL").Value

    If NIVEL < 2 Then
        MsgBox "ESTE USUÁRIO NÃO POSSUI PERMISSÃO PARA ADICIONAR OU EDITAR CADASTROS!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    

Call ATRIBUIRCAMPOS
   
    For I = 1 To 10
            CAMPO(I).BackColor = &H80000005
            
            Select Case I
            Case Is = 1, 2, 3, 4, 5, 6, 7, 8
            
            If CAMPO(I).Value = "" Then
                CAMPO(I).BackColor = &HC0C0FF
                CAMPOSOB = True
            End If
            
            End Select
    Next
        
        If CAMPOSOB = True Then
                MsgBox "PREENCHA OS CAMPOS OBRIGATÓRIOS!", vbCritical, "ATENÇÃO"
                Exit Sub
        End If
     
'CHAMADA DE FUNÇÕES/PROCEDIMENTOS'
Call PREENCHERCADASTROS

CommandButtonLIMPAR_Click
                                 
End Sub

Private Sub CommandButtonSENHAS_Click()

UserFormSENHA.Show

End Sub

Private Sub CommandButtonUSUARIOS_Click()

MultiPageGERAL.Value = 2

Dim RS As New ADODB.Recordset
Dim SQL As String

SQL = "SELECT USUÁRIO,NÍVEL FROM USUARIO"

RS.Open SQL, BD
ListBoxUSUARIOS.Clear
Do While RS.EOF = False
    ListBoxUSUARIOS.AddItem
    ListBoxUSUARIOS.List(ListBoxUSUARIOS.ListCount - 1, 0) = RS.Fields(0).Value
    ListBoxUSUARIOS.List(ListBoxUSUARIOS.ListCount - 1, 1) = RS.Fields(1).Value
    RS.MoveNext
Loop
RS.Close

End Sub

Private Sub LISTAREGISTROS_Change()

Dim ID As Long
Dim RS As New ADODB.Recordset
Dim SQL As String

If LISTAREGISTROS.Value = Empty Then Exit Sub

ID = LISTAREGISTROS.Value

SQL = "SELECT NÚCLEO,QUADRA,LOTE,DOMICÍLIO,ENTREVISTADOR FROM CADASTROS"
SQL = SQL & " WHERE ID LIKE " & ID

RS.Open SQL, BD

TextBoxFILTRONUCLEO.Value = RS.Fields(0).Value
TextBoxFILTROQUADRA.Value = RS.Fields(1).Value
TextBoxFILTROLOTE.Value = RS.Fields(2).Value
TextBoxFILTRODOMICILIO.Value = RS.Fields(3).Value
TextBoxFILTROENTREVISTADOR.Value = RS.Fields(4).Value

RS.Close

End Sub

Private Sub LISTAREGISTROS_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim ID As Long
Dim RS As New ADODB.Recordset
Dim SQL As String
Dim I As Integer
MultiPageGERAL.Value = 0

ID = LISTAREGISTROS.Value

SQL = "SELECT * FROM CADASTROS "
SQL = SQL & " WHERE ID LIKE " & ID
RS.Open SQL, BD

    For I = 1 To 81
        CAMPO(I).Value = RS.Fields(I).Value
        CAMPO(I).BackColor = &H80000005
    Next

LBMODO.Caption = "EDITAR": LBID.Caption = ID

End Sub

Private Sub MultiPageGERAL_Change()

    If MultiPageGERAL.Value = 0 Then
        FrameBUTTONS.Visible = True
    End If
    
    If MultiPageGERAL.Value = 1 Then
        FrameBUTTONS.Visible = False
    End If
    
    If MultiPageGERAL.Value = 2 Then
        FrameBUTTONS.Visible = False
        CommandButtonUSUARIOS_Click
    End If

End Sub

Private Sub TextBox1VISITA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub TextBox1VISITA_Change()

Dim DATA As String, DATA2 As String, DATA3 As String
Dim I As Integer, J As Integer, N As Integer

    DATA = TextBox1VISITA.Value
    TextBox1VISITA.MaxLength = 10
    I = Len(DATA)

    For J = 1 To I
        If IsNumeric(Mid(DATA, J, 1)) Then
            DATA2 = DATA2 & Mid(DATA, J, 1)
        End If
    Next

    I = Len(DATA2)
    
    For J = 1 To I
        DATA3 = DATA3 & Mid(DATA2, J, 1)
        If J = 3 Or J = 5 Then
            N = Len(DATA3) - 1
            DATA3 = Left(DATA3, N) & "/" & Right(DATA3, 1)
        End If
    Next

    TextBox1VISITA = DATA3

End Sub

Private Sub TextBox2VISITA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub TextBox2VISITA_Change()

Dim DATA As String, DATA2 As String, DATA3 As String
Dim I As Integer, J As Integer, N As Integer

    DATA = TextBox2VISITA.Value
    TextBox2VISITA.MaxLength = 10
    I = Len(DATA)

    For J = 1 To I
        If IsNumeric(Mid(DATA, J, 1)) Then
            DATA2 = DATA2 & Mid(DATA, J, 1)
        End If
    Next

    I = Len(DATA2)
    
    For J = 1 To I
        DATA3 = DATA3 & Mid(DATA2, J, 1)
        If J = 3 Or J = 5 Then
            N = Len(DATA3) - 1
            DATA3 = Left(DATA3, N) & "/" & Right(DATA3, 1)
        End If
    Next

    TextBox2VISITA = DATA3

End Sub

Private Sub TextBox3VISITA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub TextBox3VISITA_Change()

Dim DATA As String, DATA2 As String, DATA3 As String
Dim I As Integer, J As Integer, N As Integer

    DATA = TextBox3VISITA.Value
    TextBox3VISITA.MaxLength = 10
    I = Len(DATA)

    For J = 1 To I
        If IsNumeric(Mid(DATA, J, 1)) Then
            DATA2 = DATA2 & Mid(DATA, J, 1)
        End If
    Next

    I = Len(DATA2)

    For J = 1 To I
        DATA3 = DATA3 & Mid(DATA2, J, 1)
        If J = 3 Or J = 5 Then
            N = Len(DATA3) - 1
            DATA3 = Left(DATA3, N) & "/" & Right(DATA3, 1)
        End If
    Next

    TextBox3VISITA = DATA3

End Sub

Private Sub TextBoxCASAMENTO1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub TextBoxCASAMENTO1_Change()

Dim DATA As String, DATA2 As String, DATA3 As String
Dim I As Integer, J As Integer, N As Integer

    DATA = TextBoxCASAMENTO1.Value
    TextBoxCASAMENTO1.MaxLength = 10
    I = Len(DATA)

    For J = 1 To I
        If IsNumeric(Mid(DATA, J, 1)) Then
            DATA2 = DATA2 & Mid(DATA, J, 1)
        End If
    Next

    I = Len(DATA2)
    
    For J = 1 To I
        DATA3 = DATA3 & Mid(DATA2, J, 1)
        If J = 3 Or J = 5 Then
        N = Len(DATA3) - 1
            DATA3 = Left(DATA3, N) & "/" & Right(DATA3, 1)
        End If
    Next

    TextBoxCASAMENTO1 = DATA3

End Sub

Private Sub TextBoxCASAMENTO2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub TextBoxCASAMENTO2_Change()

Dim DATA As String, DATA2 As String, DATA3 As String
Dim I As Integer, J As Integer, N As Integer

    DATA = TextBoxCASAMENTO2.Value
    TextBoxCASAMENTO2.MaxLength = 10
    I = Len(DATA)

    For J = 1 To I
        If IsNumeric(Mid(DATA, J, 1)) Then
            DATA2 = DATA2 & Mid(DATA, J, 1)
        End If
    Next

    I = Len(DATA2)
    
    For J = 1 To I
        DATA3 = DATA3 & Mid(DATA2, J, 1)
        If J = 3 Or J = 5 Then
        N = Len(DATA3) - 1
            DATA3 = Left(DATA3, N) & "/" & Right(DATA3, 1)
        End If
    Next

    TextBoxCASAMENTO2 = DATA3

End Sub

Private Sub TextBoxCPF1_Change()

Dim CPF As String, CPF2 As String, CPF3 As String
Dim I As Integer, J As Integer, N As Integer

    CPF = TextBoxCPF1.Value
    TextBoxCPF1.MaxLength = 14
    I = Len(CPF)

    For J = 1 To I
        If IsNumeric(Mid(CPF, J, 1)) Then
            CPF2 = CPF2 & Mid(CPF, J, 1)
        End If
    Next

    I = Len(CPF2)
    
    For J = 1 To I
        CPF3 = CPF3 & Mid(CPF2, J, 1)
            If J = 4 Or J = 7 Then
            N = Len(CPF3) - 1
            CPF3 = Left(CPF3, N) & "." & Right(CPF3, 1)
        ElseIf J = 10 Then
            N = Len(CPF3) - 1
            CPF3 = Left(CPF3, N) & "-" & Right(CPF3, 1)
        End If
    Next

    TextBoxCPF1 = CPF3

End Sub

Private Sub TextBoxCPF2_Change()

Dim CPF As String, CPF2 As String, CPF3 As String
Dim I As Integer, J As Integer, N As Integer

    CPF = TextBoxCPF2.Value
    TextBoxCPF2.MaxLength = 14
    I = Len(CPF)

    For J = 1 To I
        If IsNumeric(Mid(CPF, J, 1)) Then
            CPF2 = CPF2 & Mid(CPF, J, 1)
        End If
    Next

    I = Len(CPF2)
    
    For J = 1 To I
        CPF3 = CPF3 & Mid(CPF2, J, 1)
            If J = 4 Or J = 7 Then
            N = Len(CPF3) - 1
            CPF3 = Left(CPF3, N) & "." & Right(CPF3, 1)
        ElseIf J = 10 Then
            N = Len(CPF3) - 1
            CPF3 = Left(CPF3, N) & "-" & Right(CPF3, 1)
        End If
    Next

    TextBoxCPF2 = CPF3

End Sub

Private Sub TextBoxEDIFICACOES_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    TextBoxEDIFICACOES.MaxLength = 2

End Sub

Private Sub TextBoxFILTRODOMICILIO_AfterUpdate()

Call FILTROAVANÇADO

End Sub

Private Sub TextBoxFILTROENTREVISTADOR_AfterUpdate()

Call FILTROAVANÇADO

End Sub

Private Sub TextBoxFILTROLOTE_AfterUpdate()

Call FILTROAVANÇADO

End Sub

Private Sub TextBoxFILTRONUCLEO_AfterUpdate()

Call FILTROAVANÇADO

End Sub


Private Sub TextBoxFILTROQUADRA_AfterUpdate()

Call FILTROAVANÇADO

End Sub

Private Sub TextBoxNASC1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub TextBoxNASC1_Change()

Dim DATA As String, DATA2 As String, DATA3 As String
Dim I As Integer, J As Integer, N As Integer

    DATA = TextBoxNASC1.Value
    TextBoxNASC1.MaxLength = 10
    I = Len(DATA)

    For J = 1 To I
        If IsNumeric(Mid(DATA, J, 1)) Then
        DATA2 = DATA2 & Mid(DATA, J, 1)
        End If
    Next

    I = Len(DATA2)
    
    For J = 1 To I
        DATA3 = DATA3 & Mid(DATA2, J, 1)
        If J = 3 Or J = 5 Then
        N = Len(DATA3) - 1
        DATA3 = Left(DATA3, N) & "/" & Right(DATA3, 1)
        End If
    Next

    TextBoxNASC1 = DATA3

End Sub

Private Sub TextBoxNASC2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub TextBoxNASC2_Change()

Dim DATA As String, DATA2 As String, DATA3 As String
Dim I As Integer, J As Integer, N As Integer

    DATA = TextBoxNASC2.Value
    TextBoxNASC2.MaxLength = 10
    I = Len(DATA)

    For J = 1 To I
        If IsNumeric(Mid(DATA, J, 1)) Then
        DATA2 = DATA2 & Mid(DATA, J, 1)
        End If
    Next

       I = Len(DATA2)
       
    For J = 1 To I
        DATA3 = DATA3 & Mid(DATA2, J, 1)
        If J = 3 Or J = 5 Then
        N = Len(DATA3) - 1
        DATA3 = Left(DATA3, N) & "/" & Right(DATA3, 1)
        End If
    Next

    TextBoxNASC2 = DATA3

End Sub

Private Sub TextBoxTELEFONE_AfterUpdate()

Select Case True
    Case Len(TextBoxTELEFONE) = 10
        TextBoxTELEFONE = Format(TextBoxTELEFONE, "(##) ####-####")
    Case Len(TextBoxTELEFONE) = 11
        TextBoxTELEFONE = Format(TextBoxTELEFONE, "(##) #####-####")
End Select

End Sub

Private Sub TextBoxTELEFONE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    TextBoxTELEFONE.MaxLength = 15
    
End Sub

Private Sub TextBoxTEMPOMORADIA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
    TextBoxTEMPOMORADIA.MaxLength = 2

End Sub

Private Sub UserForm_Initialize()

MultiPageGERAL.Value = 0
Call CARREGARCOMBOBOX
Call ATRIBUIRCAMPOS
Call ABRIRCONEXAO
LBID.Caption = CADASTROS.Range("ID").Value

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
       
On Error Resume Next
BD.Close
       
    If Workbooks.Count = 1 Then
        Application.Visible = True
        ThisWorkbook.Save
        Application.Quit
    Else
        Application.Visible = True
        ThisWorkbook.Close True
    End If

End Sub
