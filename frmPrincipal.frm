VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Xlsx"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar PDF"
      Height          =   495
      Left            =   6720
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7858
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin VB.DriveListBox drive 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.FileListBox file 
      Height          =   1650
      Left            =   2040
      Pattern         =   "*.xlsx"
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.DirListBox dir 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdCarregarArquivo 
      Caption         =   "Carregar"
      Height          =   480
      Left            =   5280
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblArquivo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivo"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nLinha As Long
Private nColuna As Long
Private arquivo As String
Private xlApp As Object
Private xlWorkbook As Object
Private xlWorksheet As Object
Private nomeTabela As String
Private colunas As String

Private Sub dir_Change()
    file = dir
End Sub

Private Sub drive_Change()
    dir = drive
End Sub

Private Sub Form_Load()
    Conecta
    drive = App.Path
    dir = App.Path
End Sub

Private Sub cmdCarregarArquivo_Click()
    
    grid.Clear
    arquivo = dir & "\" & file
        
    If arquivo = dir & "\" Then GoTo jump
    
    LerArquivo
    CriaTabela
    PreencheTabelaBanco
    FecharArquivo
    MsgBox "Tabela preenchida!", vbOKOnly, "Sucesso"
jump:
    
End Sub

Private Sub CriaTabela()
    Dim sql As String
    colunas = " "
    
    nomeTabela = Left(file, InStrRev(file, ".") - 1)
    cn.Execute "DROP TABLE IF EXISTS " & Replace(nomeTabela, " ", "_") & ";"
    sql = "create table IF NOT EXISTS " & Replace(nomeTabela, " ", "_") & " ( id serial primary key "
    
    With xlWorksheet.UsedRange
        nLinha = .Rows.Count
        nColuna = .Columns.Count
        grid.Cols = nColuna + 1
        ReDim Data(1 To .Rows.Count, 1 To .Columns.Count)
        grid.TextMatrix(0, 0) = "ID"
        For i = 1 To nColuna
            sql = sql & ", " & Replace(.cells(1, i).Value, " ", "_") & " VARCHAR(255) "
            If i = nColuna Then
                colunas = colunas & Replace(.cells(1, i).Value, " ", "_")
            Else
                colunas = colunas & Replace(.cells(1, i).Value, " ", "_") & ", "
            End If
            grid.TextMatrix(0, i) = Replace(.cells(1, i).Value, " ", "_")
        Next i
    End With
    
    sql = sql & ");"
    'Debug.Print sql
    cn.Execute sql
End Sub

Private Sub LerArquivo()
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Open(arquivo)
    Set xlWorksheet = xlWorkbook.Worksheets(1)
End Sub

Private Sub FecharArquivo()
    xlWorkbook.Close
    xlApp.Quit
    
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
End Sub

Private Sub PreencheTabelaBanco()
    Dim sql As String
    grid.Rows = nLinha
    With xlWorksheet.UsedRange
        ReDim Data(1 To .Rows.Count, 1 To .Columns.Count)
        For i = 2 To nLinha
            sql = "insert into " & Replace(nomeTabela, " ", "_") & " (" & colunas & ") values ("
            grid.TextMatrix(i - 1, 0) = i - 1
            For j = 1 To nColuna
                If i <= UBound(Data, 1) And j <= UBound(Data, 2) Then
                    If j = nColuna Then
                        sql = sql & "'" & .cells(i, j).Value & "') "
                    Else
                        sql = sql & "'" & .cells(i, j).Value & "', "
                    End If
                    grid.TextMatrix(i - 1, j) = .cells(i, j).Value
                End If
            Next j
            'Debug.Print sql
            cn.Execute sql
        Next i
    End With
End Sub

Private Sub cmdExportar_Click()
    frmPrint.grid.Width = (nColuna + 1) * 975
    frmPrint.grid.Height = nLinha * 255
    frmPrint.Width = frmPrint.grid.Width
    frmPrint.Height = frmPrint.grid.Height
    frmPrint.grid.Rows = grid.Rows
    frmPrint.grid.Cols = grid.Cols
    For i = 0 To grid.Rows - 1
       For j = 0 To grid.Cols - 1
           frmPrint.grid.TextMatrix(i, j) = grid.TextMatrix(i, j)
       Next j
    Next i
    
    frmPrint.PrintForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    cn.Execute "drop table " & Replace(nomeTabela, " ", "_")
End Sub

