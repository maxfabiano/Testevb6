VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Repository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private BancoConf As Object
Dim conectionStr As String

Public Sub ExecuteSQL(sql As String)
Set BancoConf = New Banco_Conf
conectionStr = BancoConf.conectionString
Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Adjust your connection string
    conn.CursorLocation = adUseClient ' Set the cursor location to client-side for a bookmarkable recordset
    conn.Open

    On Error GoTo ErrorHandler
    conn.Execute sql
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao executar SQL: " & Err.Description
End Sub
Public Sub CreateTables()
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
    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Adjust your connection string
    conn.CursorLocation = adUseClient ' Set the cursor location to client-side for a bookmarkable recordset
    conn.Open
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
Public Sub CreateStoredProcedure()
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




Public Sub CreateView()
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

Public Sub CreateFunctionCategorizarTransacao()
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


