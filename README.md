![GitHub repo size](https://img.shields.io/github/repo-size/felipebacelo/ApplicationVBA?style=for-the-badge)
![GitHub language count](https://img.shields.io/github/languages/count/felipebacelo/ApplicationVBA?style=for-the-badge)
![GitHub forks](https://img.shields.io/github/forks/felipebacelo/ApplicationVBA?style=for-the-badge)
![Bitbucket open pull requests](https://img.shields.io/bitbucket/pr-raw/felipebacelo/ApplicationVBA?style=for-the-badge)
![Bitbucket open issues](https://img.shields.io/bitbucket/issues/felipebacelo/ApplicationVBA?style=for-the-badge)

# ApplicationVBA
Aplicação em VBA Excel com Banco de Dados Access SQL

A aplicação foi desenvolvida a partir do modelo de cadastro físico do programa Cidade Legal, seguindo o conteúdo padrão do mesmo, a partir do conceito CRUD (Create, Read, Update, Delete), que representa em acrônimo as quatro operações básicas utilizadas em bases de dados relacionais fornecidas aos utilizadores do sistema.

### Desenvolvimento

Desenvolvido em Microsoft VBA Excel com banco de dados em Microsoft Access SQL.
***
### Requisitos

* Habilitar Macros
* Habilitar Guia de Desenvolvedor

### Referências às Bibliotecas

* Visual Basic For Applications
* Microsoft Excel 16.0 Object Library
* OLE Automation
* Microsoft Office 16.0 Object Library
* Microsoft Forms 2.0 Object Library
* Microsoft ActiveX Data Objects 6.1 Library
* Ref Edit Control

### Compatibilidade

Esta aplicação foi desenvolvida no Excel 2019 (64 bits) e testado no Excel 2013 (64 bits). Sua compatibilidade é garantida para a versão 2013 e superiores. Sua utilização em versões anteriores pode ocasionar em não funcionamento do mesmo.

### Usabilidade

Para utilizar o sistema de cadastro o usuário deverá:

* Realizar o download do arquivo ZIP: __ApplicationVBA__.
* Salvar o arquivo __SISTEMA DE CADASTRO.xlsm__ e __BD SISTEMA DE CADASTRO.accdb__ na mesma pasta de trabalho.

### Usuários e Senhas

Os usuários e senhas pré-definidas são:

* ADM
* ADM1 – 12345
* ADM2 – 12345
* ADM3 – 12345
***
### Passo a Passo

1º Passo - Abrir o arquivo __SISTEMA DE CADASTRO.xlsm__ com usuário e senha:

![LOGIN](https://github.com/felipebacelo/Sistema_Cadastro/blob/master/IMAGENS/LOGIN.png)

2º Passo - Após realizar o login, a tela a ser exibida será a seguinte:

![CADASTRAR](https://github.com/felipebacelo/Sistema_Cadastro/blob/master/IMAGENS/CADASTRAR.png)

Nesta tela o usuário possui acesso a todas as informações disponíveis, de acordo com seu nível de usuário, determinado pelo usuário ADM.

3º Passo - Além de cadastrar o usuário também conseguirá realizar consulta a itens já cadastrados anteriormente, através de um filtro avançado:

![CONSULTAR](https://github.com/felipebacelo/Sistema_Cadastro/blob/master/IMAGENS/CONSULTAR.png)

4º Passo - O usuário poderá redefinir sua senha atual a qualquer momento, através do Gerenciador de Senhas:

![SENHAS](https://github.com/felipebacelo/Sistema_Cadastro/blob/master/IMAGENS/SENHAS.png)

5º Passo - O usuário nível 4 é o único com permissão para adicionar, editar e deletar demais usuários no sistema através do Gerenciador de Usuários:

![USUÁRIOS](https://github.com/felipebacelo/Sistema_Cadastro/blob/master/IMAGENS/USU%C3%81RIOS.png)

***
### Exemplos de Macros Utilizadas

Macro utilizada para conexão com banco de dados Microsoft Access SQL:
```vba
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
```

Macro utilizada para edição dos registros salvos no banco de dados Microsoft Access SQL:
```vba
Sub EDITARREGISTROS(ID As Long, TODASCOLUNAS As String, REGISTRO() As String)

Dim SQL As String
Dim COLUNA() As String
Dim I As Integer
Dim STRINGFINAL As String
Dim RS As New ADODB.Recordset

COLUNA = Split(TODASCOLUNAS, ",")

For I = 1 To 81
    STRINGFINAL = STRINGFINAL & COLUNA(I - 1) & "=" & REGISTRO(I)
    If I < 81 Then STRINGFINAL = STRINGFINAL & ","
Next

STRINGFINAL = "SET " & STRINGFINAL
SQL = "Update CADASTROS " & STRINGFINAL
SQL = SQL & " WHERE ID LIKE " & ID

RS.Open SQL, BD

MsgBox "CADASTRO EDITADO COM SUCESSO!", vbInformation, "INFORMAÇÃO"

End Sub
```
***
### Licenças

_MIT License_
_Copyright   ©   2020 Felipe Bacelo Rodrigues_
