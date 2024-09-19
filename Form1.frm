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
Private TelaInicialAplication As Object
Private Sub Cadastrar_Cliente_Click()
Call TelaInicialAplication.CadastrarClienteClick
End Sub
Private Sub consultar_Click()
    Call TelaInicialAplication.consultarClick(DataGrid1, DtaInicial, DtaFinal)
End Sub
Private Sub DtaFinal_Change()
Call TelaInicialAplication.DtaFinalChange(DtaFinal)
End Sub
Private Sub DtaInicial_Change()
Call TelaInicialAplication.DtaInicialChange(DtaInicial)
End Sub
Private Sub editar_Click()
Call TelaInicialAplication.editarClick
End Sub
Private Sub excluir_Click()
    Call TelaInicialAplication.excluirClick
End Sub

Private Sub FData_Transacao_Change()
Call TelaInicialAplication.FDataTransacaoChange(FData_Transacao)
End Sub
Private Sub FDescricao_Change()
Call TelaInicialAplication.FDescricaoChange(FDescricao)
End Sub

Private Sub Form_Load()
Set TelaInicialAplication = New TelaInicialAplication
Call TelaInicialAplication.FormLoad(DataGrid1, FNumero_Cartao)
End Sub

Private Sub FValor_Transacao_Change()
Call TelaInicialAplication.FValorTransacaoChange(FValor_Transacao)
End Sub
Private Sub inserir_Click()
    Call TelaInicialAplication.inserirClick(DataGrid1, FNumero_Cartao, FValor_Transacao, FData_Transacao)
End Sub
Private Sub salvar_Click()
Call TelaInicialAplication.salvarClick
End Sub

Private Sub TFuncao_Click()
Call TelaInicialAplication.TFuncaoClick(DataGrid1)
End Sub

Private Sub TTransacoes_Click()
Call TelaInicialAplication.TTransacoesClick(DataGrid1, DtaInicial, DtaFinal)
End Sub
Private Sub TView_Click()
Call TelaInicialAplication.TViewClick(DataGrid1, DtaFinal, DtaInicial)
End Sub
