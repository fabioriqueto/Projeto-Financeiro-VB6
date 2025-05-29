VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form CLIENTES 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes                                                                *** Megatron***"
   ClientHeight    =   5010
   ClientLeft      =   2190
   ClientTop       =   2190
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7575
   Begin VB.TextBox Text12 
      DataField       =   "CPF_CNPJ"
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
      Left            =   1200
      TabIndex        =   39
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      DataField       =   "RG_IE"
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
      Left            =   1200
      TabIndex        =   37
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      DataField       =   "OBS"
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
      Height          =   765
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox Text9 
      DataField       =   "TEL2"
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
      Left            =   1800
      TabIndex        =   33
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      DataField       =   "TEL1"
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
      TabIndex        =   32
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      DataField       =   "CEP"
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
      Left            =   3120
      TabIndex        =   30
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "UF"
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
      Height          =   315
      ItemData        =   "CLIENTES.frx":0000
      Left            =   360
      List            =   "CLIENTES.frx":0055
      OLEDropMode     =   1  'Manual
      TabIndex        =   28
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      DataField       =   "CIDADE"
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
      Left            =   2280
      TabIndex        =   26
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      DataField       =   "BAIRRO"
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
      TabIndex        =   24
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "END"
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
      Left            =   1080
      TabIndex        =   22
      Top             =   720
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Localizações"
      Height          =   2415
      Left            =   3480
      TabIndex        =   17
      Top             =   2520
      Width           =   3975
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Sistema\Bancod01.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CLIENTES"
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "CLIENTES.frx":0190
         Height          =   1695
         Left            =   120
         OleObjectBlob   =   "CLIENTES.frx":01A0
         TabIndex        =   18
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Indexadores"
      Height          =   1095
      Left            =   4560
      TabIndex        =   12
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton Option4 
         Caption         =   "CEP"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "CPF / CNPJ"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "RG / I. Estadual"
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      DataField       =   "NOME"
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
      Left            =   1080
      TabIndex        =   11
      Top             =   360
      Width           =   3375
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
      Left            =   2640
      TabIndex        =   10
      Top             =   4560
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
      Left            =   1800
      TabIndex        =   9
      Top             =   4560
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
      Left            =   960
      TabIndex        =   8
      Top             =   4560
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
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Editar"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Novo"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
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
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "CPF/CNPJ:"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "RG/IE:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Observações:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Telefones:"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "CEP:"
      Height          =   255
      Left            =   2760
      TabIndex        =   29
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "UF:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Cidade"
      Height          =   255
      Left            =   2280
      TabIndex        =   25
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Bairro:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Código:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "CLIENTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Text2.Enabled = True
DBGrid1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command2.Enabled = False
Command1.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Text3.Enabled = False
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Combo1.Enabled = True
Frame1.Enabled = False
End Sub

Private Sub Command2_Click()
Data1.Recordset.Edit
Text2.Enabled = True
DBGrid1.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command2.Enabled = False
Command1.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Text3.Enabled = False
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Combo1.Enabled = True
Frame1.Enabled = False
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
DBGrid1.Enabled = True
Text3.Enabled = True
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Combo1.Enabled = False
Frame1.Enabled = True
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
DBGrid1.Enabled = True
Text3.Enabled = True
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Combo1.Enabled = False
Frame1.Enabled = True
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

Private Sub Form_Activate()
Set db1 = OpenDatabase("c:\sistema\bancod01.mdb")
Set rs1 = db1.OpenRecordset("CLIENTES")
Set Data1.Recordset = rs1
Data1.Recordset.Index = "INOME"
Label3.Caption = "Digite o nome:"


End Sub

Private Sub Option1_Click()
Label3.Caption = "Digite o nome:"
Data1.Recordset.Index = "INOME"
End Sub

Private Sub Option2_Click()
Label3.Caption = "Digite o RG/IE:"
Data1.Recordset.Index = "IRG_IE"
End Sub

Private Sub Option3_Click()
Label3.Caption = "Digite o CPF/CNPJ:"
Data1.Recordset.Index = "ICPF_CNPJ"
End Sub

Private Sub Option4_Click()
Label3.Caption = "Digite o CEP:"
Data1.Recordset.Index = "ICEP"
End Sub

Private Sub Text3_Change()
Data1.Recordset.Seek ">=", Text3.Text
End Sub
