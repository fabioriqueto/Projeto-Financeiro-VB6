VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Financeiras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Financeiras   *Megatron*"
   ClientHeight    =   3510
   ClientLeft      =   3045
   ClientTop       =   1185
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3840
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FINANCEIRO"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "TAXATROCA"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   1440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   327680
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "%#,##0.000;(%#,##0.000)"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command8 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Editar"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Novo"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      DataField       =   "FINANCEIRO"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      DataField       =   "CODIGO"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Financeiras.frx":0000
      Height          =   1095
      Left            =   120
      OleObjectBlob   =   "Financeiras.frx":0010
      TabIndex        =   12
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Taxa de Juros:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Financeira:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Financeiras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Text2.Enabled = True
MaskEdBox1.Enabled = True
DBGrid1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command2.Enabled = False
Command1.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
End Sub

Private Sub Command2_Click()
Data1.Recordset.Edit
Text2.Enabled = True
MaskEdBox1.Enabled = True
DBGrid1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command2.Enabled = False
Command1.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False

End Sub

Private Sub Command3_Click()
Data1.Recordset.Update
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Text2.Enabled = False
MaskEdBox1.Enabled = False
DBGrid1.Enabled = True
Data1.Refresh
End Sub

Private Sub Command4_Click()
Data1.Recordset.CancelUpdate
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Text2.Enabled = False
MaskEdBox1.Enabled = False
DBGrid1.Enabled = True
End Sub

Private Sub Command5_Click()
Data1.Recordset.MoveFirst
End Sub

Private Sub Command6_Click()
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF Then Data1.Recordset.MoveNext
End Sub

Private Sub Command7_Click()
Data1.Recordset.MoveNext
If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious
End Sub

Private Sub Command8_Click()
Data1.Recordset.MoveLast
End Sub

