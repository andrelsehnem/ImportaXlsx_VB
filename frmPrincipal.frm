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
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7858
      _Version        =   393216
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
      Height          =   360
      Left            =   5280
      TabIndex        =   1
      Top             =   1800
      Width           =   990
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

Private Sub cmdCarregarArquivo_Click()
    arquivo = dir & "\" & file
    LerArquivo
    CriaTabela
End Sub

Private Sub dir_Change()
    file = dir
End Sub

Private Sub drive_Change()
    dir = drive
End Sub

Private Sub Form_Load()
    Dim b As Boolean
    
    Conecta
    drive = App.Path
    dir = App.Path
    
End Sub

Private Sub CriaTabela()
    Dim posBarra As Integer
    Dim nomeTabela As String
    Dim sql As String
    
    nomeTabela = file
    sql = "create table " & nomeTabela & " ( id serial primary key "
    
    With xlWorksheet.UsedRange
        nLinha = .Rows.Count
        nColuna = .Columns.Count
        ReDim data(1 To .Rows.Count, 1 To .Columns.Count)
        For i = 1 To nLinha
            sql = sql & ", " & .Cells(1, i).Value & " VARCHAR(255) "
            Debug.Print sql
        Next i
    End With
    
    sql = sql & ");"
    Debug.Print (sql);
    cn.Execute sql
End Sub


Private Sub LerArquivo()
    Dim data() As Variant
    Dim i As Long, j As Long

    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Open(arquivo)
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    
End Sub

Private Sub LerDados()
    With xlWorksheet.UsedRange
        ReDim data(1 To .Rows.Count, 1 To .Columns.Count)
        For i = 1 To nLinha
            For j = 1 To nColuna
                If i <= UBound(data, 1) And j <= UBound(data, 2) Then
                    data(i, j) = .Cells(i, j).Value
                End If
            Next j
        Next i
    End With
    
    xlWorkbook.Close
    xlApp.Quit
    
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    
    For i = 1 To UBound(data, 1)
        For j = 1 To UBound(data, 2)
            Debug.Print data(i, j) & " "
        Next j
    Next i
End Sub
