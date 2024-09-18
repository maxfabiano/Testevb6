VERSION 5.00
Begin VB.Form frmEditarTransacao 
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12135
   LinkTopic       =   "Editar"
   ScaleHeight     =   5955
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ENumeroCartao 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   8055
   End
   Begin VB.TextBox EData 
      Height          =   525
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   6495
   End
   Begin VB.TextBox EValor 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   4800
      Width           =   11175
   End
   Begin VB.ComboBox lstTransacoes 
      Height          =   315
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   12135
   End
   Begin VB.Label label4 
      Caption         =   "Data Da Transação"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Valor Da Transação"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Numero Do Cartão"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Pesquisar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmEditarTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExecuteSQL(sql As String)
    On Error GoTo ErrorHandler
    conn.Execute sql
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao executar SQL: " & Err.Description
End Sub



Private Sub Command1_Click()
Dim sql As String
    Dim cmd As ADODB.Command
    Dim Id_Transacao As String
    Dim Numero_Cartao As String
    Dim Valor_Transacao As String
    Dim Data_Transacao As String
    Dim selectedItem As String
    If Len(Trim(EValor.Text)) = 0 Then
        EValor.Text = "1,01"
    End If
    If Len(Trim(EData.Text)) < 10 Then
        EData.Text = Format(Now, "dd/mm/yyyy")
    End If
    If CartaoExiste(ENumeroCartao.Text) Then
        MsgBox "Cartão encontrado no sistema.", vbInformation
        Exit Sub
    End If
        
    selectedItem = lstTransacoes.List(lstTransacoes.ListIndex)
    ' Assuming these values are parsed from the selected item in your ComboBox or ListBox
    Id_Transacao = Split(selectedItem, " - ")(0)
    Numero_Cartao = ENumeroCartao.Text
    Valor_Transacao = EValor.Text
    Data_Transacao = EData.Text
    
    
    ' Prepare the SQL Update query
    sql = "UPDATE Transacoes SET Numero_Cartao = ?, Valor_Transacao = ?, Data_Transacao = ? WHERE Id_Transacao = ?"
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.Open
    ' Create the Command object
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn  ' Assuming conn is your open ADODB.Connection object
        .CommandText = sql
        .CommandType = adCmdText
        
        ' Add parameters for the query
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, Numero_Cartao)
        .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 50, Replace(Valor_Transacao, ",", "."))
        .Parameters.Append .CreateParameter(, adDate, adParamInput, , Format(Data_Transacao, "yyyy-mm-dd")) ' Ensure date format for MySQL
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , CLng(Id_Transacao)) ' Id_Transacao should be numeric
        
        ' Execute the update query
        .Execute
    End With
    
    MsgBox "Transação atualizada com sucesso!"
    
End Sub

Private Sub EData_Change()
Dim texto As String
    Dim dataFormatada As String
    EData.Text = RemoveNonNumeric(EData.Text)
    ' Remove todos os caracteres que não são números
    texto = Replace(EData.Text, "/", "")
    
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
        EData.Text = dataFormatada
        
        ' Mantém o cursor no final do texto
        EData.SelStart = Len(EData.Text)
    End If
End Sub
Private Function CartaoExiste(Numero_Cartao As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim sql As String

    ' SQL query to check if the card number exists in the Clientes table
    sql = "SELECT Numero_Cartao FROM my_database.Clientes WHERE Numero_Cartao = '" & Numero_Cartao & "'"
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.Open
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

Private Sub Cartoes()
Dim rs As ADODB.Recordset
    Dim sql As String

    ' Initialize and open the connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.Open

    ' Clear the ListBox before populating
    ENumeroCartao.Clear
    
    ' Define the SQL query to get all the transactions
    sql = "SELECT Numero_Cartao, Nome_Cliente FROM Clientes LIMIT 10"
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.Open
    ' Open the Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient ' Important for static cursor
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    ' Populate the ListBox with the records
    Do Until rs.EOF
        ENumeroCartao.AddItem rs.Fields("Numero_Cartao").Value
                              
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

Private Sub ENumeroCartao_Change()
Dim rs As ADODB.Recordset
    Dim sql As String
    Dim searchTerm As String
    ENumeroCartao.Text = RemoveNonNumeric(ENumeroCartao.Text)
    If Len(Trim(ENumeroCartao.Text)) = 0 Then
        Call Cartoes
        Exit Sub
    End If
    If Len(Trim(ENumeroCartao.Text)) < 3 Then
        
        Exit Sub
    End If
    ' Get the search term
    searchTerm = ENumeroCartao.Text
    ' Clear the ListBox
    Dim i As Integer

    ' Loop through the ListBox and remove each item starting from the last one
    For i = ENumeroCartao.ListCount - 1 To 0 Step -1
        ENumeroCartao.RemoveItem i
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
        ENumeroCartao.AddItem rs.Fields("Numero_Cartao").Value
        rs.MoveNext
    Loop
End Sub

Private Sub EValor_Change()
Dim texto As String
    Dim valor As Double
    EValor.Text = RemoveNonNumeric(EValor.Text)
    ' Remove caracteres não numéricos temporariamente
    texto = Replace(EValor.Text, ",", "")
    texto = Replace(texto, ".", "")
    
    ' Verifica se o campo não está vazio
    If Len(texto) > 0 Then
        ' Converte a string para um valor numérico e divide por 100 para obter as casas decimais
        valor = CDbl(texto) / 100
        
        ' Formata o valor como moeda (com vírgula e duas casas decimais)
        EValor.Text = Format(valor, "###0.00")
        
        ' Mantém o cursor no final do texto
        EValor.SelStart = Len(EValor.Text)
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim sql As String

    ' Initialize and open the connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.Open

    ' Clear the ListBox before populating
    lstTransacoes.Clear
    
    ' Define the SQL query to get all the transactions
    sql = "SELECT Id_Transacao, Numero_Cartao, Valor_Transacao,Data_Transacao FROM Transacoes LIMIT 10"
    
    ' Open the Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient ' Important for static cursor
    rs.Open sql, conn, adOpenStatic, adLockReadOnly

    ' Populate the ListBox with the records
    Do Until rs.EOF
        lstTransacoes.AddItem rs.Fields("Id_Transacao").Value & " - Numero Cartao: " & _
                              rs.Fields("Numero_Cartao").Value & " - Data Da Transacao: " & _
                              rs.Fields("Data_Transacao").Value & " - Valor Da Transacao: " & _
                              Format(rs.Fields("Valor_Transacao").Value, "#,##0.00")
                              
        rs.MoveNext
    Loop
    
    ' Close the recordset
    rs.Close
    Set rs = Nothing
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


Private Sub lstTransacoes_Change()
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim searchTerm As String
    ' Get the search term
    searchTerm = lstTransacoes.Text
    ' Clear the ListBox
    Dim i As Integer

    ' Loop through the ListBox and remove each item starting from the last one
    For i = lstTransacoes.ListCount - 1 To 0 Step -1
        lstTransacoes.RemoveItem i
    Next i
    
    ' Define the SQL query to search for transactions based on the search term
    sql = "SELECT Id_Transacao, Numero_Cartao, Valor_Transacao,Data_Transacao FROM Clientes " & _
          "WHERE Numero_Cartao LIKE '%" & searchTerm & "%' OR " & _
          "Valor_Transacao LIKE '%" & searchTerm & "%' LIMIT 10"
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DSN=odbc1;UID=root;PWD=root_password;DATABASE=my_database;" ' Include the database name here
    conn.Open
    ' Open the Recordset
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Populate the ListBox with the search results
    Do Until rs.EOF
        lstTransacoes.AddItem rs.Fields("Id_Transacao").Value & " - Numero Cartao: " & _
                              rs.Fields("Numero_Cartao").Value & " - Data Da Transacao: " & _
                              rs.Fields("Data_Transacao").Value & " - Valor Da Transacao: " & _
                              Format(rs.Fields("Valor_Transacao").Value, "#,##0.00")
        rs.MoveNext
    Loop
End Sub





Private Sub lstTransacoes_Click()
Dim selectedItem As String
selectedItem = lstTransacoes.List(lstTransacoes.ListIndex)
    ENumeroCartao.Text = Split(Split(selectedItem, " - ")(1), ":")(1)
    EValor.Text = Split(Split(selectedItem, " - ")(3), ":")(1)
    EData.Text = Split(Split(selectedItem, " - ")(2), ":")(1)
End Sub
