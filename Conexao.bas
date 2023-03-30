Attribute VB_Name = "Conexao"
Public strCon As String
Public cn As ADODB.Connection
'strCon = "Driver={PostgreSQL ANSI};Server=localhostPort=" & frmPrincipal.txtPorta_Cli.Text & ";Database=bancoModelo;Uid=postgres;Pwd=admin;sslmode=disable;"

Public Function Conecta() As Boolean
    strCon = "Driver={PostgreSQL ANSI};Server=localhost;Database=bancoModelo;Uid=postgres;Pwd=admin;sslmode=disable;"
    cn.CursorLocation = adUseServer
    cn.ConnectionString = st
    Debug.Print strCon
    On Error GoTo ErroConex
        Tj.Open
        Debug.Print Tj.State
        Conectado = True
        Exit Function
ErroConex:
            MsgBox "Erero ao Conectar"
            Debug.Print Error
            Conectado = False
            If cn.State = 1 Then
                cn.Close
            End If
            Exit Function
            
End Function



' https://www.macoratti.net/08/02/vb_cdcli.htm

' https://macoratti.net/10/10/vb_xls2.htm

' https://macoratti.net/08/02/vb_cdcli.htm

' https://macoratti.net/vb6_msfg.htm

' https://macoratti.net/vb6_exp.htm

' https://macoratti.net/vb6grids.htm
