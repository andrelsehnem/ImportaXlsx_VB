Attribute VB_Name = "Conexao"
Public strCon As String
Public cn As ADODB.Connection

Public Function Conecta() As Boolean
    strCon = "Driver={PostgreSQL Unicode};Server=localhost;Database=bancoModelo;Uid=postgres;Pwd=admin"
    
    Set cn = New ADODB.Connection
    cn.CursorLocation = adUseServer
    cn.ConnectionString = strCon
    Debug.Print strCon
    On Error GoTo ErroConex
    cn.Open
    Debug.Print cn.State
    Conectado = True
    Exit Function
    
ErroConex:
    MsgBox "Erro ao Conectar"
    Debug.Print Error
    Conectado = False
End Function

