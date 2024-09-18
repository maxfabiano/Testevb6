VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormTelaInicial 
   Caption         =   "Tela Inicial"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   15150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton TFuncao 
      Caption         =   "Função"
      Height          =   495
      Left            =   4920
      TabIndex        =   23
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox FNumero_Cartao 
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton TTransacoes 
      Caption         =   "Tabela Transações"
      Height          =   495
      Left            =   8280
      TabIndex        =   21
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton TView 
      Caption         =   "View"
      Height          =   495
      Left            =   6360
      TabIndex        =   20
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox FDescricao 
      Height          =   975
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox DtaFinal 
      Height          =   375
      Left            =   11640
      TabIndex        =   14
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox DtaInicial 
      Height          =   375
      Left            =   10200
      TabIndex        =   13
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton salvar 
      Caption         =   "Exportar Excel"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7920
      Width           =   14655
   End
   Begin VB.CommandButton Cadastrar_Cliente 
      Caption         =   "Cadastrar Cliente"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TextBoxSQL 
      Height          =   195
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox FData_Transacao 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox FValor_Transacao 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$"" #.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4575
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   8070
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton consultar 
      Caption         =   "consultar"
      Height          =   375
      Left            =   13080
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton excluir 
      Caption         =   "Excluir Transação"
      Height          =   495
      Left            =   12480
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton editar 
      Caption         =   "Editar Transação"
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton inserir 
      BackColor       =   &H8000000D&
      Caption         =   "Inserir Transação"
      Height          =   375
      Left            =   6840
      MaskColor       =   &H00C0C000&
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Tabela Transaçoes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Descrição"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Dta Final"
      Height          =   255
      Left            =   11640
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Dta Inicial"
      Height          =   255
      Left            =   10200
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label_Data_Transacao 
      Caption         =   "Data Transação"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label_Valor_Transacao 
      Caption         =   "Valor Transacao"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label_Numero_Cartao 
      Caption         =   "Numero Cartao"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "FormTelaInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declare variables for database connection
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim Id_Transacao As String
Dim Numero_Cartao As String
Dim Valor_Transacao As String
Dim Data_Transacao As String
Dim Descricao As String
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long


Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Sub Cadastrar_Cliente_Click()
FormCliente.Show vbModal ' Exibe o Form2 e bloqueia o Form1 até que o Form2 seja fechado
End Sub
Private Sub CriarUsuarioAdmin()
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim numeroCartao As String
    Dim nomeCliente As String
    Dim emailCliente As String
    Dim telefoneCliente As String
    Dim dataAtual As String
    
    ' Definindo os valores do usuário administrador
    numeroCartao = "123456"
    nomeCliente = "Administrador"
    emailCliente = "admin@exemplo.com"
    telefoneCliente = "123456789"
    
    ' Obtendo a data atual no formato AAAA-MM-DD
    dataAtual = Format(Date, "yyyy-mm-dd")
    
    ' Verifica se o cliente com o número do cartão já existe
    sql = "SELECT * FROM Clientes WHERE Numero_Cartao = '" & numeroCartao & "'"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Se não encontrar registros, insere o novo cliente
    If rs.EOF Then
        sql = "INSERT INTO Clientes (Numero_Cartao, Nome_Cliente, Email, Telefone) " & _
              "VALUES ('" & numeroCartao & "', '" & nomeCliente & "', '" & emailCliente & "', '" & telefoneCliente & "')"
        conn.Execute sql
        
    Else
        
    End If
    
    ' Fechar o Recordset
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Command2_Click()

End Sub

Private Sub consultar_Click()
    Dim sql As String
    Dim cmd As ADODB.Command


    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Adjust your connection string
    conn.CursorLocation = adUseClient ' Set the cursor location to client-side for a bookmarkable recordset
    conn.Open
    ' Create and configure the Command object
        
        ' Add the parameters for the query
        If Len(Trim(DtaInicial.Text)) = 0 Then
        DtaInicial.Text = "01011900" ' Set to 1st January 1900 if empty
    Else
        ' Convert the date to MySQL format yyyy-mm-dd
        
    End If

    ' Check if DtaFinal.Text is empty, if so, set it to the current date
    If Len(Trim(DtaFinal.Text)) = 0 Then
        DtaFinal.Text = Format(Now, "ddmmyyyy") ' Set to current date if empty
    Else
        ' Convert the date to MySQL format yyyy-mm-dd
        
    End If
    
    dataOriginalInicial = DtaInicial.Text
    
    ' Usa a função Split para separar a data nas barras "/"
    partesDataInicial = Split(dataOriginalInicial, "/")
    
    ' Reorganiza a data no formato AAAA-MM-DD (formato MySQL)
    dataFormatadaInicial = partesDataInicial(2) & "-" & partesDataInicial(1) & "-" & partesDataInicial(0)
    
    
    
    
    
    
    dataOriginalFinal = DtaFinal.Text
    
    ' Usa a função Split para separar a data nas barras "/"
    partesDataFinal = Split(dataOriginalFinal, "/")
    
    ' Reorganiza a data no formato AAAA-MM-DD (formato MySQL)
    dataFormatadaFinal = partesDataFinal(2) & "-" & partesDataFinal(1) & "-" & partesDataFinal(0)
    


Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandText = "call sp_TotalTransacoesPorPeriodo(?, ?)"
    cmd.CommandType = adCmdText

    ' Add parameters for the stored procedure
    cmd.Parameters.Append cmd.CreateParameter("Data_Inicial", adDate, adParamInput, , dataFormatadaInicial)
    cmd.Parameters.Append cmd.CreateParameter("Data_Final", adDate, adParamInput, , dataFormatadaFinal)

    ' Execute the stored procedure and get the recordset
    Set rs = cmd.Execute

    ' Check if there are any records
    If rs.EOF Then
        MsgBox "No transactions found for the specified period."
        Exit Sub
    End If

    ' Bind the recordset to the DataGrid
    Set DataGrid1.DataSource = rs
End Sub




Private Sub DtaFinal_Change()
Dim texto As String
    Dim dataFormatada As String
    DtaFinal.Text = RemoveNonNumeric(DtaFinal.Text)
    ' Remove todos os caracteres que não são números
    texto = Replace(DtaFinal.Text, "/", "")
    
    ' Verifica se o campo não está vazio
    If Len(texto) > 0 Then
        ' Limita o texto a 8 caracteres (DDMMAAAA)
        If Len(texto) > 8 Then
            texto = Left(texto, 8)
        End If
        
        ' Formata a data conforme a quantidade de dígitos digitados
        Select Case Len(texto)
            Case 1 To 2 ' Dia
                dataFormatada = texto
            Case 3 To 4 ' Dia/Mês
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2)
            Case 5 To 8 ' Dia/Mês/Ano
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2) & "/" & Mid(texto, 5)
        End Select
        
        ' Atualiza o campo de texto com a data formatada
        DtaFinal.Text = dataFormatada
        
        ' Mantém o cursor no final do texto
        DtaFinal.SelStart = Len(DtaFinal.Text)
    End If
End Sub

Private Sub DtaInicial_Change()
Dim texto As String
    Dim dataFormatada As String
    DtaInicial.Text = RemoveNonNumeric(DtaInicial.Text)
    ' Remove todos os caracteres que não são números
    texto = Replace(DtaInicial.Text, "/", "")
    
    ' Verifica se o campo não está vazio
    If Len(texto) > 0 Then
        ' Limita o texto a 8 caracteres (DDMMAAAA)
        If Len(texto) > 8 Then
            texto = Left(texto, 8)
        End If
        
        ' Formata a data conforme a quantidade de dígitos digitados
        Select Case Len(texto)
            Case 1 To 2 ' Dia
                dataFormatada = texto
            Case 3 To 4 ' Dia/Mês
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2)
            Case 5 To 8 ' Dia/Mês/Ano
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2) & "/" & Mid(texto, 5)
        End Select
        
        ' Atualiza o campo de texto com a data formatada
        DtaInicial.Text = dataFormatada
        
        ' Mantém o cursor no final do texto
        DtaInicial.SelStart = Len(DtaInicial.Text)
    End If
End Sub

Private Sub editar_Click()
frmEditarTransacao.Show
End Sub

Private Sub excluir_Click()
    frmExcluirTransacao.Show
End Sub

Private Sub FData_Transacao_Change()
Dim texto As String
    Dim dataFormatada As String
    FData_Transacao.Text = RemoveNonNumeric(FData_Transacao.Text)
    ' Remove todos os caracteres que não são números
    texto = Replace(FData_Transacao.Text, "/", "")
    
    ' Verifica se o campo não está vazio
    If Len(texto) > 0 Then
        ' Limita o texto a 8 caracteres (DDMMAAAA)
        If Len(texto) > 8 Then
            texto = Left(texto, 8)
        End If
        
        ' Formata a data conforme a quantidade de dígitos digitados
        Select Case Len(texto)
            Case 1 To 2 ' Dia
                dataFormatada = texto
            Case 3 To 4 ' Dia/Mês
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2)
            Case 5 To 8 ' Dia/Mês/Ano
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2) & "/" & Mid(texto, 5)
        End Select
        
        ' Atualiza o campo de texto com a data formatada
        FData_Transacao.Text = dataFormatada
        
        ' Mantém o cursor no final do texto
        FData_Transacao.SelStart = Len(FData_Transacao.Text)
    End If
End Sub


Private Sub FDescricao_Change()
FDescricao.Text = RemoveSpecialCharacters(FDescricao.Text)
End Sub
Private Function RemoveSpecialCharacters(inputString As String) As String
    Dim i As Integer
    Dim cleanedString As String
    cleanedString = ""

    ' Loop through each character in the input string
    For i = 1 To Len(inputString)
        ' Check if the character is a number or a letter (upper or lower case)
        If (Mid(inputString, i, 1) >= "0" And Mid(inputString, i, 1) <= "9") Or _
           (Mid(inputString, i, 1) >= "A" And Mid(inputString, i, 1) <= "Z") Or _
           (Mid(inputString, i, 1) >= "a" And Mid(inputString, i, 1) <= "z") Then
            cleanedString = cleanedString & Mid(inputString, i, 1)
        End If
    Next i

    ' Return the cleaned string that contains only numbers and letters
    RemoveSpecialCharacters = cleanedString
End Function

Private Sub FNumero_Cartao_Change()
Dim rs As ADODB.Recordset
    Dim sql As String
    Dim searchTerm As String
    FNumero_Cartao.Text = RemoveNonNumeric(FNumero_Cartao.Text)
    If Len(Trim(FNumero_Cartao.Text)) = 0 Then
        Call Cartoes
        Exit Sub
    End If
    If Len(Trim(FNumero_Cartao.Text)) < 3 Then
        
        Exit Sub
    End If
    ' Get the search term
    searchTerm = FNumero_Cartao.Text
    ' Clear the ListBox
    Dim i As Integer

    ' Loop through the ListBox and remove each item starting from the last one
    For i = FNumero_Cartao.ListCount - 1 To 0 Step -1
        FNumero_Cartao.RemoveItem i
    Next i
    
    ' Define the SQL query to search for transactions based on the search term
    sql = "SELECT Numero_Cartao FROM Clientes " & _
          "WHERE Numero_Cartao LIKE '%" & searchTerm & "%'"
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.CursorLocation = adUseClient
    conn.Open
    ' Open the Recordset
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Populate the ListBox with the search results
    Do Until rs.EOF
        FNumero_Cartao.AddItem rs.Fields("Numero_Cartao").Value
        rs.MoveNext
    Loop
End Sub

Private Sub Form_Load()
    ' Initialize the connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;"
    conn.Open
    ' Cria todas as tabelas e funções necessárias
    Call CreateTables
    Call CreateFunctionCategorizarTransacao
    Call CreateStoredProcedure
    Call CreateView
    
    ' Inicializa os campos de entrada
    ResetInputFields
    Call ConsultarTransacoes
    Call Cartoes
End Sub


Private Sub Cartoes()
Dim rs As ADODB.Recordset
    Dim sql As String

    ' Initialize and open the connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.Open

    ' Clear the ListBox before populating
    FNumero_Cartao.Clear
    
    ' Define the SQL query to get all the transactions
    sql = "SELECT Numero_Cartao, Nome_Cliente FROM Clientes LIMIT 10"
    
    ' Open the Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient ' Important for static cursor
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    ' Populate the ListBox with the records
    Do Until rs.EOF
        FNumero_Cartao.AddItem rs.Fields("Numero_Cartao").Value
                              
        rs.MoveNext
    Loop
    
    ' Close the recordset
    rs.Close
End Sub


Private Sub ResetInputFields()
    Id_Transacao = ""
    Numero_Cartao = ""
    Valor_Transacao = 0 ' Inicializa como 0, pois é um valor numérico
    Data_Transacao = ""
    Descricao = "" ' Inicializa como string vazia
End Sub

Private Function CartaoExiste(Numero_Cartao As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim sql As String

    ' SQL query to check if the card number exists in the Clientes table
    sql = "SELECT Numero_Cartao FROM my_database.Clientes WHERE Numero_Cartao = '" & Numero_Cartao & "'"
    
    ' Initialize recordset and execute the query
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    ' Check if a record was returned (card number exists)
    If Not rs.EOF Then
        CartaoExiste = False
    Else
        CartaoExiste = True
    End If
    
    ' Close the recordset
    rs.Close
    Set rs = Nothing
End Function


Private Sub ConverterDataParaMySQL()
    Dim dataOriginal As String
    Dim partesData() As String
    Dim dataFormatada As String

    ' Supondo que a data no formato DD/MM/AAAA esteja na variável dataOriginal
    dataOriginal = FData_Transacao.Text
    
    ' Usa a função Split para separar a data nas barras "/"
    partesData = Split(dataOriginal, "/")
    
    ' Reorganiza a data no formato AAAA-MM-DD (formato MySQL)
    dataFormatada = partesData(2) & "-" & partesData(1) & "-" & partesData(0)
    
    Data_Transacao = dataFormatada
End Sub

Private Sub carregarStrings()
    If Len(Trim(FValor_Transacao.Text)) = 0 Then
        FValor_Transacao.Text = "1,01"
    End If
    ' Carrega os valores dos campos de texto
    Numero_Cartao = FNumero_Cartao.Text
    Valor_Transacao = Replace(FValor_Transacao.Text, ",", ".")
    If Len(Trim(FData_Transacao.Text)) < 10 Then
        FData_Transacao.Text = Format(Now, "dd/mm/yyyy")
    End If
    Call ConverterDataParaMySQL
End Sub


Private Sub CreateStoredProcedure()
    Dim sql As String
    ' Exclui a procedure se já existir
    sql = "DROP PROCEDURE IF EXISTS sp_TotalTransacoesPorPeriodo;"
    Call ExecuteSQL(sql)
    
    ' Cria a nova stored procedure
    sql = "CREATE PROCEDURE sp_TotalTransacoesPorPeriodo(" & _
          "IN Data_Inicial DATE, IN Data_Final DATE) " & _
          "BEGIN " & _
          "SELECT Numero_Cartao, SUM(Valor_Transacao) AS Valor_Total, COUNT(*) AS Quantidade_Transacoes " & _
          "FROM Transacoes " & _
          "WHERE Data_Transacao BETWEEN Data_Inicial AND Data_Final " & _
          "GROUP BY Numero_Cartao; " & _
          "END;"
    Call ExecuteSQL(sql)
End Sub

Private Sub CreateTables()
    Dim sql As String
    sql = "CREATE DATABASE IF NOT EXISTS my_database"

    ' Executa o SQL no servidor MySQL
    Call ExecuteSQL(sql)

    ' Define o banco de dados a ser utilizado
    sql = "USE my_database"
    Call ExecuteSQL(sql)
    ' Criação da tabela Clientes
    sql = "CREATE TABLE IF NOT EXISTS Clientes (" & _
          "Numero_Cartao INT AUTO_INCREMENT PRIMARY KEY, " & _
          "Nome_Cliente VARCHAR(100) NOT NULL, " & _
          "Email VARCHAR(100), " & _
          "Telefone VARCHAR(20)" & _
          ");"
    Call ExecuteSQL(sql)

    ' Criação da tabela Transacoes
    sql = "CREATE TABLE IF NOT EXISTS Transacoes (" & _
      "Id_Transacao INT AUTO_INCREMENT PRIMARY KEY, " & _
      "Numero_Cartao INT, " & _
      "Valor_Transacao DECIMAL(10,2) NOT NULL, " & _
      "Data_Transacao DATE NOT NULL, " & _
      "Descricao VARCHAR(255), " & _
      "Status INT, " & _
      "FOREIGN KEY (Numero_Cartao) REFERENCES Clientes(Numero_Cartao) ON DELETE CASCADE" & _
      ");"
        Call ExecuteSQL(sql)


    ' Criação da tabela Categorias
    sql = "CREATE TABLE IF NOT EXISTS Categorias (" & _
          "Id_Categoria INT AUTO_INCREMENT PRIMARY KEY, " & _
          "Descricao_Categoria VARCHAR(100) NOT NULL" & _
          ");"
    Call ExecuteSQL(sql)
    
    sql = "INSERT INTO Categorias (Id_Categoria, Descricao_Categoria) " & _
      "SELECT 1, 'Alta' FROM DUAL WHERE NOT EXISTS (SELECT 1 FROM Categorias WHERE Descricao_Categoria = 'Alta') UNION ALL " & _
      "SELECT 2, 'Média' FROM DUAL WHERE NOT EXISTS (SELECT 1 FROM Categorias WHERE Descricao_Categoria = 'Média') UNION ALL " & _
      "SELECT 3, 'Baixa' FROM DUAL WHERE NOT EXISTS (SELECT 1 FROM Categorias WHERE Descricao_Categoria = 'Baixa');"
Call ExecuteSQL(sql)

    ' Criação da tabela Transacoes_Categorias
    sql = "CREATE TABLE IF NOT EXISTS Transacoes_Categorias (" & _
          "Id INT AUTO_INCREMENT PRIMARY KEY, " & _
          "Id_Transacao INT, " & _
          "Id_Categoria INT, " & _
          "FOREIGN KEY (Id_Transacao) REFERENCES Transacoes(Id_Transacao) ON DELETE CASCADE, " & _
          "FOREIGN KEY (Id_Categoria) REFERENCES Categorias(Id_Categoria) ON DELETE CASCADE" & _
          ");"
    Call ExecuteSQL(sql)

    ' Criação da tabela Auditoria
    sql = "CREATE TABLE IF NOT EXISTS Auditoria (" & _
          "Id_Auditoria INT AUTO_INCREMENT PRIMARY KEY, " & _
          "Descricao VARCHAR(255), " & _
          "Data_Auditoria TIMESTAMP DEFAULT CURRENT_TIMESTAMP, " & _
          "Id_Transacao INT, " & _
          "FOREIGN KEY (Id_Transacao) REFERENCES Transacoes(Id_Transacao) ON DELETE SET NULL" & _
          ");"
    Call ExecuteSQL(sql)
    Call CriarUsuarioAdmin
End Sub


Private Sub CreateView()
    Dim sql As String
    
    ' Exclui a view se já existir
    sql = "DROP VIEW IF EXISTS vw_TransacoesComCategoria;"
    Call ExecuteSQL(sql)
    
    ' Cria a nova view
    sql = "CREATE VIEW vw_TransacoesComCategoria AS " & _
          "SELECT c.Nome_Cliente, t.Numero_Cartao, t.Valor_Transacao, t.Data_Transacao, " & _
          "cat.Descricao_Categoria AS Categoria " & _
          "FROM Transacoes t " & _
          "JOIN Clientes c ON t.Numero_Cartao = c.Numero_Cartao " & _
          "LEFT JOIN Transacoes_Categorias tc ON t.Id_Transacao = tc.Id_Transacao " & _
          "LEFT JOIN Categorias cat ON tc.Id_Categoria = cat.Id_Categoria;"
    Call ExecuteSQL(sql)
End Sub

Private Sub CreateFunctionCategorizarTransacao()
    Dim sql As String

    ' Exclui a função se já existir
    sql = "DROP FUNCTION IF EXISTS CategorizarTransacao;"
    Call ExecuteSQL(sql)

    ' Cria a função CategorizarTransacao
    sql = "CREATE FUNCTION CategorizarTransacao(Valor DECIMAL(10,2)) " & _
          "RETURNS VARCHAR(10) " & _
          "DETERMINISTIC " & _
          "BEGIN " & _
          "   IF Valor > 1000 THEN " & _
          "       RETURN 'Alta'; " & _
          "   ELSEIF Valor BETWEEN 500 AND 1000 THEN " & _
          "       RETURN 'Média'; " & _
          "   ELSE " & _
          "       RETURN 'Baixa'; " & _
          "   END IF; " & _
          "END;"
    Call ExecuteSQL(sql)
End Sub

Private Sub ExecuteSQL(sql As String)
    On Error GoTo ErrorHandler
    conn.Execute sql
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao executar SQL: " & Err.Description
End Sub







Private Sub ConsultarTransacoes()
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Set the connection's cursor location to client-side to support bookmarks
    conn.CursorLocation = adUseClient
    
    ' Define your SQL query
    sql = "SELECT * FROM Transacoes where Status=1"
    
    ' Open the recordset with a bookmarkable cursor type
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Bind the recordset to the DataGrid
    Set DataGrid1.DataSource = rs
    
    DataGrid1.AllowUpdate = False

    ' Optionally lock individual columns
    Dim i As Integer
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).Locked = True
    Next i

End Sub





Private Sub FValor_Transacao_Change()
Dim texto As String
    Dim valor As Double
    FValor_Transacao.Text = RemoveNonNumeric(FValor_Transacao.Text)
    ' Remove caracteres não numéricos temporariamente
    texto = Replace(FValor_Transacao.Text, ",", "")
    texto = Replace(texto, ".", "")
    
    ' Verifica se o campo não está vazio
    If Len(texto) > 0 Then
        ' Converte a string para um valor numérico e divide por 100 para obter as casas decimais
        valor = CDbl(texto) / 100
        
        ' Formata o valor como moeda (com vírgula e duas casas decimais)
        FValor_Transacao.Text = Format(valor, "###0.00")
        
        ' Mantém o cursor no final do texto
        FValor_Transacao.SelStart = Len(FValor_Transacao.Text)
    End If
End Sub

Private Function RemoveNonNumeric(inputString As String) As String
    Dim i As Integer
    Dim cleanedString As String
    cleanedString = ""

    ' Loop through each character in the input string
    For i = 1 To Len(inputString)
        ' Check if the character is numeric
        If Mid(inputString, i, 1) >= "0" And Mid(inputString, i, 1) <= "9" Then
            cleanedString = cleanedString & Mid(inputString, i, 1)
        End If
    Next i

    ' Return the cleaned string that contains only numbers
    RemoveNonNumeric = cleanedString
End Function


Private Sub inserir_Click()
    Call carregarStrings
    If CartaoExiste(FNumero_Cartao.Text) Then
        MsgBox "Cartão encontrado no sistema.", vbInformation
        Exit Sub
    End If
    ' Monta a string SQL para inserir a transação
    Dim sql As String
    sql = "INSERT INTO Transacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status) " & _
          "VALUES ('" & Numero_Cartao & "', " & Valor_Transacao & ", '" & Data_Transacao & "', '" & FDescricao & "', 1)"

    ' Executa a inserção da transação
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = adCmdText
        ' Executa o comando para inserir a transação
        .Execute
    End With
    
    ' Pega o ID da transação recém inserida
    Dim rs As ADODB.Recordset
    Set rs = conn.Execute("SELECT LAST_INSERT_ID() AS Id_Transacao")
    Dim Id_Transacao As Long
    Id_Transacao = rs.Fields("Id_Transacao").Value
    rs.Close
    
    ' Determina a categoria com base no Valor_Transacao
    Dim Id_Categoria As Integer
    If Valor_Transacao > 1000 Then
        Id_Categoria = 1 ' Categoria "Alta"
    ElseIf Valor_Transacao >= 500 And Valor_Transacao <= 1000 Then
        Id_Categoria = 2 ' Categoria "Média"
    Else
        Id_Categoria = 3 ' Categoria "Baixa"
    End If

    ' Insere o registro na tabela Transacoes_Categorias
    sql = "INSERT INTO Transacoes_Categorias (Id_Transacao, Id_Categoria) " & _
          "VALUES (" & Id_Transacao & ", " & Id_Categoria & ")"
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = adCmdText
        ' Executa o comando para inserir na tabela Transacoes_Categorias
        .Execute
    End With

    MsgBox "Transação e categoria inseridas com sucesso!"
    Call ConsultarTransacoes
End Sub


Private Sub salvar_Click()
Dim rs As ADODB.Recordset
    Dim sql As String
    Dim linha As String
    Dim saveFilePath As String
    Dim fileNum As Integer
    
    ' Query to get transactions from the last month
    sql = "SELECT Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, CategorizarTransacao(Valor_Transacao) AS Categoria " & _
          "FROM Transacoes " & _
          "WHERE Status = 1 AND Data_Transacao >= DATE_SUB(CURDATE(), INTERVAL 1 MONTH)"
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.CursorLocation = adUseClient ' Set the cursor location to client-side for a bookmarkable recordset
    conn.Open
    ' Create a new ADODB.Recordset object and execute the query
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    TextBoxSQL.Text = sql
    ' Open a Save File Dialog to let the user choose where to save the file
    saveFilePath = ShowSaveFileDialog("xls") ' Call a function to show SaveFileDialog for CSV
    
    If saveFilePath = "" Then
        MsgBox "Exportação cancelada."
        Exit Sub
    End If
    
    ' Get a file number and open the file for output
    fileNum = FreeFile
    Open saveFilePath For Output As #fileNum
    
    ' Write the header line
    linha = "Numero Cartao,Valor Transacao,Data Transacao,Descricao,Categoria"
    Print #fileNum, linha
    
    ' Write the data from the recordset
    Do Until rs.EOF
        linha = rs("Numero_Cartao") & "," & _
                Replace(CStr(rs("Valor_Transacao")), ",", ".") & "," & _
                rs("Data_Transacao") & "," & _
                rs("Descricao") & "," & _
                rs("Categoria")
        Print #fileNum, linha
        rs.MoveNext
    Loop
    
    ' Close the file
    Close #fileNum
    
    ' Optionally, you can rename the file to have an .xls extension
    Name saveFilePath As Replace(saveFilePath, saveFilePath, saveFilePath + ".xls")
    
    MsgBox "Relatório exportado com sucesso para: " & saveFilePath
End Sub

Private Sub TFuncao_Click()
Dim sql As String
    Dim rs As ADODB.Recordset

    ' Query to fetch the transaction details and the category using the function CategorizarTransacao
    sql = "SELECT Id_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, " & _
          "CategorizarTransacao(Valor_Transacao) AS Categoria " & _
          "FROM Transacoes"
    
    ' Initialize and open the connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.CursorLocation = adUseClient ' Set the cursor location to client-side for a bookmarkable recordset
    conn.Open
    
    ' Create and open the Recordset with a bookmarkable cursor type
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient ' Client-side cursor
    rs.Open sql, conn, adOpenStatic, adLockReadOnly ' Use adOpenStatic for bookmarkable recordset

    ' Bind the Recordset to the DataGrid
    Set DataGrid1.DataSource = rs
End Sub

Private Function ShowSaveFileDialog(fileExtension As String) As String
    Dim ofn As OPENFILENAME
    Dim lRet As Long
    Dim sFilter As String
    Dim sFile As String
    
    ' Initialize structure
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hWnd
    sFilter = "CSV Files (*." & fileExtension & ")" & Chr(0) & "*." & fileExtension & Chr(0)
    ofn.lpstrFilter = sFilter
    ofn.nFilterIndex = 1
    sFile = String(255, 0)
    ofn.lpstrFile = sFile
    ofn.nMaxFile = Len(sFile)
    ofn.Flags = OFN_OVERWRITEPROMPT
    ofn.lpstrTitle = "Salvar Relatório"
    
    ' Show dialog
    lRet = GetSaveFileName(ofn)
    
    ' Return the selected file path
    If lRet <> 0 Then
        ShowSaveFileDialog = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
    Else
        ShowSaveFileDialog = ""
    End If
End Function

Private Sub TTransacoes_Click()
Dim sql As String
    Dim cmd As ADODB.Command

    ' Construct the SQL query to filter by Numero_Cartao and date range
    sql = "SELECT * FROM Transacoes WHERE Data_Transacao >= ? AND Data_Transacao <= ? "
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.CursorLocation = adUseClient ' Set the cursor location to client-side for a bookmarkable recordset
    conn.Open
    ' Create and configure the Command object
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = adCmdText
        
        ' Add the parameters for the query
        If Len(Trim(DtaInicial.Text)) = 0 Then
        DtaInicial.Text = "01011900" ' Set to 1st January 1900 if empty
    Else
        ' Convert the date to MySQL format yyyy-mm-dd
        
    End If

    ' Check if DtaFinal.Text is empty, if so, set it to the current date
    If Len(Trim(DtaFinal.Text)) = 0 Then
        DtaFinal.Text = "01019900" ' Set to current date if empty
    Else
        ' Convert the date to MySQL format yyyy-mm-dd
        
    End If

        ' Filter by Data_Transacao range (between DtaInicial.Text and DtaFinal.Text)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , Format(DtaInicial.Text, "yyyy-mm-dd"))
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , Format(DtaFinal.Text, "yyyy-mm-dd"))
        
        ' Execute the query and set the result to the Recordset (rs)
        Set rs = .Execute
    End With

    ' Bind the result to the DataGrid
    Set DataGrid1.DataSource = rs
End Sub

Private Sub TView_Click()
Dim sql As String
    Dim cmd As ADODB.Command

    ' Construct the SQL query to filter by Numero_Cartao and date range
    sql = "SELECT * FROM vw_TransacoesComCategoria WHERE Data_Transacao >= ? AND Data_Transacao <= ? "
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.CursorLocation = adUseClient ' Set the cursor location to client-side for a bookmarkable recordset
    conn.Open
    ' Create and configure the Command object
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = adCmdText
        
        ' Add the parameters for the query
        If Len(Trim(DtaInicial.Text)) = 0 Then
        DtaInicial.Text = "01011900" ' Set to 1st January 1900 if empty
    Else
        ' Convert the date to MySQL format yyyy-mm-dd
        
    End If

    ' Check if DtaFinal.Text is empty, if so, set it to the current date
    If Len(Trim(DtaFinal.Text)) = 0 Then
        DtaFinal.Text = "01019900" ' Set to current date if empty
    Else
        ' Convert the date to MySQL format yyyy-mm-dd
        
    End If

        ' Filter by Data_Transacao range (between DtaInicial.Text and DtaFinal.Text)
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , Format(DtaInicial.Text, "yyyy-mm-dd"))
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , Format(DtaFinal.Text, "yyyy-mm-dd"))
        
        ' Execute the query and set the result to the Recordset (rs)
        Set rs = .Execute
    End With

    ' Bind the result to the DataGrid
    Set DataGrid1.DataSource = rs
End Sub
