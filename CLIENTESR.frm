VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form CLIENTESR 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimentação de Clientes                            *** Megatron***"
   ClientHeight    =   5205
   ClientLeft      =   2190
   ClientTop       =   2190
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5895
   Begin VB.CommandButton Command3 
      Caption         =   "EXCL."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4680
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Relatórios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   4455
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "CLIENTESR.frx":0000
      Left            =   120
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Sistema\movcli01.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.Data Data3 
      Caption         =   "Filtrado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Data Data2 
      Caption         =   "Geral"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Frame Frame2 
      Caption         =   "Localizações"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton Option5 
         Caption         =   "Mov Baixa"
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Taxa dev. não pagas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Cheques devolvidos"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Mov. cheques pré"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mov. troca"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text1 
         DataField       =   "CODIGO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Sistema\Bancod01.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CLIENTES"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   2400
         Width           =   3255
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "CLIENTESR.frx":0010
         Height          =   975
         Left            =   120
         OleObjectBlob   =   "CLIENTESR.frx":0020
         TabIndex        =   1
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Data para consulta:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Digite o nome para pesquisa:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   2175
      End
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "CLIENTESR.frx":09DA
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "CLIENTESR.frx":09EA
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Filtrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   5655
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Bindings        =   "CLIENTESR.frx":298C
      Left            =   720
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Sistema\movcli02.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "CLIENTESR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim recRecordset1 As Recordset, recRecordset2 As Recordset
Set recRecordset1 = Data2.Recordset 'copy the recordset
teste2 = Date
If Option1 = True Then teste2 = CDate(MaskEdBox1)
If Option5 = True Then teste2 = CDate(MaskEdBox1)
teste2 = "cdate('" + Trim(Str(teste2)) + "')"
If Option5 = True Then teste = "ci = " + Text1.Text + " and PRE_DATA = " + teste2
If Option4 = True Then teste = "ci = " + Text1.Text + " and " + "qtd_dev <> qtd_devp"
If Option3 = True Then teste = "ci = " + Text1.Text + " and " + "qtd_dev <> 0"
If Option2 = True Then teste = "ci = " + Text1.Text + " and pre_data >= " + teste2
If Option1 = True Then teste = "ci = " + Text1.Text + " and data = " + teste2
recRecordset1.Filter = teste
Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type)
Set Data3.Recordset = recRecordset2 'assign back to original recordset object
DBGrid2.Visible = True
Command1.Visible = False
Command3.Visible = True
End Sub


Private Sub Command2_Click()
If Option1.Value = True Then CrystalReport1.PrintReport
If Option2.Value = True Then CrystalReport2.PrintReport
End Sub

Private Sub Command3_Click()
Data3.Recordset.Delete
Data3.Refresh
Data2.Refresh
Command3.Visible = False
Command1.Visible = True
DBGrid2.Visible = False
End Sub

Private Sub DBGrid1_Click()
Command1.Visible = True
DBGrid2.Visible = False
Command3.Visible = False
End Sub

Private Sub Form_Activate()
Set db1 = OpenDatabase("c:\sistema\bancod01.mdb")
Set rs1 = db1.OpenRecordset("CLIENTES")
Set Data1.Recordset = rs1
Data1.Recordset.Index = "INOME"
MaskEdBox1 = Date
Data1.Recordset.MoveFirst
End Sub

Private Sub MaskEdBox1_Change()
If MaskEdBox1.FormattedText = "" Then MaskEdBox1 = Date
Command1.Visible = True
DBGrid2.Visible = False

End Sub

Private Sub Option1_Click()
Command3.Visible = False
If Option1.Value = True Then Command2.Enabled = True
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = True
End Sub

Private Sub Option2_Click()
Command3.Visible = False
If Option2.Value = True Then Command2.Enabled = True
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = False
End Sub

Private Sub Option3_Click()
Command3.Visible = False
If Option3.Value = True Then Command2.Enabled = False
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = False
End Sub

Private Sub Option4_Click()
Command3.Visible = False
If Option4.Value = True Then Command2.Enabled = True
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = False
End Sub

Private Sub Option5_Click()
Command3.Visible = False
If Option5.Value = True Then Command2.Enabled = False
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = True

End Sub

Private Sub Text3_Change()
Data1.Recordset.Seek ">=", Text3.Text
End Sub
