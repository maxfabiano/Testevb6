VERSION 5.00
Begin VB.Form FormCliente 
   Caption         =   "Cadastrar Cliente"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10665
   LinkTopic       =   "Form2"
   ScaleHeight     =   4290
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BCadastrar 
      Caption         =   "Cadastrar"
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox FEmail 
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox FNumeroCartao 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox FNomeCLiente 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox FTelefone 
      Height          =   285
      Left            =   5640
      TabIndex        =   0
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "Cadastrar Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Telefone"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Email"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Numero Do Cartão"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Nome Cliente"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   3015
   End
End
Attribute VB_Name = "FormCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private BancoConf As Object
Dim conectionStr As String

Private Sub BCadastrar_Click()
Set BancoConf = New Banco_Conf
     conectionStr = BancoConf.conectionString
Dim rs As ADODB.Recordset
    Dim sql As String
    Dim numeroCartao As String
    Dim nomeCliente As String
    Dim emailCliente As String
    Dim telefoneCliente As String
    Dim dataAtual As String
    
    ' Definindo os valores do usuário administrador
    numeroCartao = FNumeroCartao.Text
    nomeCliente = FNomeCLiente.Text
    emailCliente = FEmail.Text
    telefoneCliente = FTelefone.Text
    
    ' Obtendo a data atual no formato AAAA-MM-DD
    dataAtual = Format(Date, "yyyy-mm-dd")
    
    ' Verifica se o cliente com o número do cartão já existe
    sql = "SELECT * FROM Clientes WHERE Numero_Cartao = '" & numeroCartao & "'"
    Set conn = New ADODB.Connection
    conn.ConnectionString = conectionStr ' Include the database name here
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
    Unload FormCliente
    MsgBox "Cadastro Realizado"
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

Private Sub FNomeCLiente_Change()
    FNomeCLiente.Text = RemoveSpecialCharacters(FNomeCLiente.Text)

End Sub

Private Sub FNumeroCartao_Change()
    FNumeroCartao.Text = RemoveNonNumeric(FNumeroCartao.Text)
End Sub

Private Sub FTelefone_Change()
FTelefone.Text = RemoveNonNumeric(FTelefone.Text)
End Sub
