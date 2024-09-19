VERSION 5.00
Begin VB.Form frmExcluirTransacao 
   Caption         =   "Excluir"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox lstTransacoes 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   600
      Width           =   11295
   End
   Begin VB.CommandButton btnExcluir 
      Caption         =   "Excluir"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Pesquisar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmExcluirTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private BancoConf As Object
Dim conectionStr As String

Private Sub btnExcluir_Click()
Dim selectedItem As String
    Dim Id_Transacao As String
    Dim sql As String
    
    ' Get the selected item from the ListBox
    selectedItem = lstTransacoes.List(lstTransacoes.ListIndex)
    
    ' Extract the Id_Transacao from the selected item (assuming it's the first field)
    Id_Transacao = Split(selectedItem, " - ")(0)
    
    ' Confirm the deletion
    If MsgBox("Você tem certeza que deseja excluir a transação " & Id_Transacao & "?", vbYesNo + vbQuestion, "Confirmar exclusão") = vbYes Then
        ' Define the SQL query to delete the selected transaction
        sql = "DELETE FROM Transacoes WHERE Id_Transacao = " & Id_Transacao
        
        Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Include the database name here
    conn.Open
        ' Execute the delete query
        conn.Execute sql
        
        ' Refresh the ListBox after deletion
        Call Form_Load ' Re-populate the ListBox
        MsgBox "Transação excluída com sucesso!"
    End If
End Sub


Private Sub ExecuteSQL(sql As String)
    On Error GoTo ErrorHandler
    conn.Execute sql
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao executar SQL: " & Err.Description
End Sub


Private Sub Form_Load()
Set BancoConf = New Banco_Conf
     conectionStr = BancoConf.conectionString
    Dim rs As ADODB.Recordset
    Dim sql As String

    ' Initialize and open the connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Include the database name here
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
    sql = "SELECT Id_Transacao, Numero_Cartao, Valor_Transacao,Data_Transacao FROM Transacoes " & _
          "WHERE Numero_Cartao LIKE '%" & searchTerm & "%' OR " & _
          "Valor_Transacao LIKE '%" & searchTerm & "%' LIMIT 10"
    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Include the database name here
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
