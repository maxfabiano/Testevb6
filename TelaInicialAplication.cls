VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TelaInicialAplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Declare variables for database connection
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim Id_Transacao As String
Dim Numero_Cartao As String
Dim Valor_Transacao As String
Dim Data_Transacao As String
Dim Descricao As String
Public BancoConf As Object
Dim conectionStr As String
Public Repositorio As Object
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

Public Sub CadastrarClienteClick()
FormCliente.Show vbModal ' Exibe o Form2 e bloqueia o Form1 at� que o Form2 seja fechado
End Sub

Public Sub consultarClick(DataGrid1 As DataGrid, DtaInicial As TextBox, DtaFinal As TextBox)
    Dim sql As String
    Dim cmd As ADODB.Command


    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Adjust your connection string
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
    
    ' Usa a fun��o Split para separar a data nas barras "/"
    partesDataInicial = Split(dataOriginalInicial, "/")
    
    ' Reorganiza a data no formato AAAA-MM-DD (formato MySQL)
    dataFormatadaInicial = partesDataInicial(2) & "-" & partesDataInicial(1) & "-" & partesDataInicial(0)
    
    
    
    
    
    
    dataOriginalFinal = DtaFinal.Text
    
    ' Usa a fun��o Split para separar a data nas barras "/"
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




Public Sub DtaFinalChange(DtaFinal As TextBox)
Dim texto As String
    Dim dataFormatada As String
    DtaFinal.Text = RemoveNonNumeric(DtaFinal.Text)
    ' Remove todos os caracteres que n�o s�o n�meros
    texto = Replace(DtaFinal.Text, "/", "")
    
    ' Verifica se o campo n�o est� vazio
    If Len(texto) > 0 Then
        ' Limita o texto a 8 caracteres (DDMMAAAA)
        If Len(texto) > 8 Then
            texto = Left(texto, 8)
        End If
        
        ' Formata a data conforme a quantidade de d�gitos digitados
        Select Case Len(texto)
            Case 1 To 2 ' Dia
                dataFormatada = texto
            Case 3 To 4 ' Dia/M�s
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2)
            Case 5 To 8 ' Dia/M�s/Ano
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2) & "/" & Mid(texto, 5)
        End Select
        
        ' Atualiza o campo de texto com a data formatada
        DtaFinal.Text = dataFormatada
        
        ' Mant�m o cursor no final do texto
        DtaFinal.SelStart = Len(DtaFinal.Text)
    End If
End Sub

Public Sub DtaInicialChange(DtaInicial As TextBox)
Dim texto As String
    Dim dataFormatada As String
    DtaInicial.Text = RemoveNonNumeric(DtaInicial.Text)
    ' Remove todos os caracteres que n�o s�o n�meros
    texto = Replace(DtaInicial.Text, "/", "")
    
    ' Verifica se o campo n�o est� vazio
    If Len(texto) > 0 Then
        ' Limita o texto a 8 caracteres (DDMMAAAA)
        If Len(texto) > 8 Then
            texto = Left(texto, 8)
        End If
        
        ' Formata a data conforme a quantidade de d�gitos digitados
        Select Case Len(texto)
            Case 1 To 2 ' Dia
                dataFormatada = texto
            Case 3 To 4 ' Dia/M�s
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2)
            Case 5 To 8 ' Dia/M�s/Ano
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2) & "/" & Mid(texto, 5)
        End Select
        
        ' Atualiza o campo de texto com a data formatada
        DtaInicial.Text = dataFormatada
        
        ' Mant�m o cursor no final do texto
        DtaInicial.SelStart = Len(DtaInicial.Text)
    End If
End Sub

Public Sub editarClick()
frmEditarTransacao.Show
End Sub

Public Sub excluirClick()
    frmExcluirTransacao.Show
End Sub

Public Sub FDataTransacaoChange(FData_Transacao As TextBox)
Dim texto As String
    Dim dataFormatada As String
    FData_Transacao.Text = RemoveNonNumeric(FData_Transacao.Text)
    ' Remove todos os caracteres que n�o s�o n�meros
    texto = Replace(FData_Transacao.Text, "/", "")
    
    ' Verifica se o campo n�o est� vazio
    If Len(texto) > 0 Then
        ' Limita o texto a 8 caracteres (DDMMAAAA)
        If Len(texto) > 8 Then
            texto = Left(texto, 8)
        End If
        
        ' Formata a data conforme a quantidade de d�gitos digitados
        Select Case Len(texto)
            Case 1 To 2 ' Dia
                dataFormatada = texto
            Case 3 To 4 ' Dia/M�s
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2)
            Case 5 To 8 ' Dia/M�s/Ano
                dataFormatada = Left(texto, 2) & "/" & Mid(texto, 3, 2) & "/" & Mid(texto, 5)
        End Select
        
        ' Atualiza o campo de texto com a data formatada
        FData_Transacao.Text = dataFormatada
        
        ' Mant�m o cursor no final do texto
        FData_Transacao.SelStart = Len(FData_Transacao.Text)
    End If
End Sub


Public Sub FDescricaoChange(FDescricao As TextBox)
FDescricao.Text = RemoveSpecialCharacters(FDescricao.Text)
End Sub
Public Function RemoveSpecialCharacters(inputString As String) As String
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

Public Sub FNumeroCartaoChange(FNumero_Cartao As ComboBox)
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
          "WHERE Numero_Cartao LIKE '%" & searchTerm & "%' LIMIT 10000"
    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Include the database name here
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

Public Sub FormLoad(DataGrid1 As DataGrid, FNumero_Cartao As ComboBox)
    Set BancoConf = New Banco_Conf
     conectionStr = BancoConf.conectionString
     Set Repositorio = New Repository
     
    ' Initialize the connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;"
    conn.Open
    ' Cria todas as tabelas e fun��es necess�rias
    Call Repositorio.CreateTables
    Call Repositorio.CreateFunctionCategorizarTransacao
    Call Repositorio.CreateStoredProcedure
    Call Repositorio.CreateView
    
    ' Inicializa os campos de entrada
    ResetInputFields
    Call ConsultarTransacoes(DataGrid1)
    Call Cartoes(FNumero_Cartao)
End Sub


Public Sub Cartoes(FNumero_Cartao As ComboBox)
Dim rs As ADODB.Recordset
    Dim sql As String

    ' Initialize and open the connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Include the database name here
    conn.Open

    ' Clear the ListBox before populating
    FNumero_Cartao.Clear
    
    ' Define the SQL query to get all the transactions
    sql = "SELECT Numero_Cartao, Nome_Cliente FROM Clientes LIMIT 10000"
    
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


Public Sub ResetInputFields()
    Id_Transacao = ""
    Numero_Cartao = ""
    Valor_Transacao = 0 ' Inicializa como 0, pois � um valor num�rico
    Data_Transacao = ""
    Descricao = "" ' Inicializa como string vazia
End Sub

Public Function CartaoExiste(Numero_Cartao As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim sql As String

    ' SQL query to check if the card number exists in the Clientes table
    sql = "SELECT Numero_Cartao FROM Clientes WHERE Numero_Cartao = '" & Numero_Cartao & "' LIMIT 10000"
    
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


Public Sub ConverterDataParaMySQL(FData_Transacao As TextBox)
    Dim dataOriginal As String
    Dim partesData() As String
    Dim dataFormatada As String

    ' Supondo que a data no formato DD/MM/AAAA esteja na vari�vel dataOriginal
    dataOriginal = FData_Transacao.Text
    
    ' Usa a fun��o Split para separar a data nas barras "/"
    partesData = Split(dataOriginal, "/")
    
    ' Reorganiza a data no formato AAAA-MM-DD (formato MySQL)
    dataFormatada = partesData(2) & "-" & partesData(1) & "-" & partesData(0)
    
    Data_Transacao = dataFormatada
End Sub

Public Sub carregarStrings(DataGrid1 As DataGrid, FValor_Transacao As TextBox, FNumero_Cartao As ComboBox, FData_Transacao As TextBox)
    If Len(Trim(FValor_Transacao.Text)) = 0 Then
        FValor_Transacao.Text = "1,01"
    End If
    ' Carrega os valores dos campos de texto
    Numero_Cartao = FNumero_Cartao.Text
    Valor_Transacao = Replace(FValor_Transacao.Text, ",", ".")
    If Len(Trim(FData_Transacao.Text)) < 10 Then
        FData_Transacao.Text = Format(Now, "dd/mm/yyyy")
    End If
    Call ConverterDataParaMySQL(FData_Transacao)
End Sub

Public Sub ConsultarTransacoes(DataGrid1 As DataGrid)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Set the connection's cursor location to client-side to support bookmarks
    conn.CursorLocation = adUseClient
    
    ' Define your SQL query
    sql = "SELECT * FROM Transacoes where Status=1 LIMIT 10000"
    
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





Public Sub FValorTransacaoChange(FValor_Transacao As TextBox)
Dim texto As String
    Dim valor As Double
    FValor_Transacao.Text = RemoveNonNumeric(FValor_Transacao.Text)
    ' Remove caracteres n�o num�ricos temporariamente
    texto = Replace(FValor_Transacao.Text, ",", "")
    texto = Replace(texto, ".", "")
    
    ' Verifica se o campo n�o est� vazio
    If Len(texto) > 0 Then
        ' Converte a string para um valor num�rico e divide por 100 para obter as casas decimais
        valor = CDbl(texto) / 100
        
        ' Formata o valor como moeda (com v�rgula e duas casas decimais)
        FValor_Transacao.Text = Format(valor, "###0.00")
        
        ' Mant�m o cursor no final do texto
        FValor_Transacao.SelStart = Len(FValor_Transacao.Text)
    End If
End Sub

Public Function RemoveNonNumeric(inputString As String) As String
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


Public Sub inserirClick(DataGrid1 As DataGrid, FNumero_Cartao As ComboBox, FValor_Transacao As TextBox, FData_Transacao As TextBox, FDescricao As TextBox)
    Call carregarStrings(DataGrid1, FValor_Transacao, FNumero_Cartao, FData_Transacao)
    If CartaoExiste(FNumero_Cartao.Text) Then
        MsgBox "Cart�o encontrado no sistema.", vbInformation
        Exit Sub
    End If
    ' Monta a string SQL para inserir a transa��o
    Dim sql As String
    sql = "INSERT INTO Transacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status) " & _
          "VALUES ('" & Numero_Cartao & "', " & Valor_Transacao & ", '" & Data_Transacao & "', '" & FDescricao & "', 1)"

    ' Executa a inser��o da transa��o
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = adCmdText
        ' Executa o comando para inserir a transa��o
        .Execute
    End With
    
    ' Pega o ID da transa��o rec�m inserida
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
        Id_Categoria = 2 ' Categoria "M�dia"
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

    MsgBox "Transa��o e categoria inseridas com sucesso!"
    Call ConsultarTransacoes(DataGrid1)
End Sub


Public Sub salvarClick()
Dim rs As ADODB.Recordset
    Dim sql As String
    Dim linha As String
    Dim saveFilePath As String
    Dim fileNum As Integer
    
    ' Query to get transactions from the last month
    sql = "SELECT Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, CategorizarTransacao(Valor_Transacao) AS Categoria " & _
          "FROM Transacoes " & _
          "WHERE Status = 1 AND Data_Transacao >= DATE_SUB(CURDATE(), INTERVAL 1 MONTH) LIMIT 10000"
    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Include the database name here
    conn.CursorLocation = adUseClient ' Set the cursor location to client-side for a bookmarkable recordset
    conn.Open
    ' Create a new ADODB.Recordset object and execute the query
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    ' Open a Save File Dialog to let the user choose where to save the file
    saveFilePath = ShowSaveFileDialog("xls") ' Call a function to show SaveFileDialog for CSV
    
    If saveFilePath = "" Then
        MsgBox "Exporta��o cancelada."
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
    
    MsgBox "Relat�rio exportado com sucesso para: " & saveFilePath
End Sub

Public Sub TFuncaoClick(DataGrid1 As DataGrid)
Dim sql As String
    Dim rs As ADODB.Recordset

    ' Query to fetch the transaction details and the category using the function CategorizarTransacao
    sql = "SELECT Id_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, " & _
          "CategorizarTransacao(Valor_Transacao) AS Categoria " & _
          "FROM Transacoes LIMIT 10000"
    
    ' Initialize and open the connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Include the database name here
    conn.CursorLocation = adUseClient ' Set the cursor location to client-side for a bookmarkable recordset
    conn.Open
    
    ' Create and open the Recordset with a bookmarkable cursor type
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient ' Client-side cursor
    rs.Open sql, conn, adOpenStatic, adLockReadOnly ' Use adOpenStatic for bookmarkable recordset

    ' Bind the Recordset to the DataGrid
    Set DataGrid1.DataSource = rs
End Sub

Public Function ShowSaveFileDialog(fileExtension As String) As String
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
    ofn.lpstrTitle = "Salvar Relat�rio"
    
    ' Show dialog
    lRet = GetSaveFileName(ofn)
    
    ' Return the selected file path
    If lRet <> 0 Then
        ShowSaveFileDialog = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
    Else
        ShowSaveFileDialog = ""
    End If
End Function

Public Sub TTransacoesClick(DataGrid1 As DataGrid, DtaInicial As TextBox, DtaFinal As TextBox)
Dim sql As String
    Dim cmd As ADODB.Command

    ' Construct the SQL query to filter by Numero_Cartao and date range
    sql = "SELECT * FROM Transacoes WHERE Data_Transacao >= ? AND Data_Transacao <= ?  LIMIT 10000"
    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Include the database name here
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

Public Sub TViewClick(DataGrid1 As DataGrid, DtaFinal As TextBox, DtaInicial As TextBox)
Dim sql As String
    Dim cmd As ADODB.Command

    ' Construct the SQL query to filter by Numero_Cartao and date range
    sql = "SELECT * FROM vw_TransacoesComCategoria WHERE Data_Transacao >= ? AND Data_Transacao <= ? LIMIT 10000"
    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Include the database name here
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


