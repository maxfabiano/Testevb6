VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Banco_Conf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function conectionString() As String
    ' Define e retorna a string de conex�o usada em todo codigo
    Dim uid As String
    Dim pwd As String
    Dim database As String
    uid = "root"
    pwd = "root_password"
    database = "my_database"
    
    conectionString = "DSN=odbc1;UID=" + uid + ";PWD=" + pwd + ";DATABASE=" + database + ";"
End Function

