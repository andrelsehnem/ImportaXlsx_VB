VERSION 5.00
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
   Begin VB.CommandButton cmdCarregarArquivo 
      Caption         =   "Carregar"
      Height          =   360
      Left            =   8760
      TabIndex        =   3
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton cmdSelecionarArquivo 
      Caption         =   "Selecionar"
      Height          =   360
      Left            =   7560
      TabIndex        =   2
      Top             =   240
      Width           =   990
   End
   Begin VB.TextBox txtCaminhoArquivo 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   360
      Width           =   7215
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


Private Sub Form_Load()
    Dim b As Boolean
    
    Conecta
    
End Sub
