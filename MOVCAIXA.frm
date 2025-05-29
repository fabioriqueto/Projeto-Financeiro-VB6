VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form MOVCAIXA 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimentação de Caixa                               *** Megatron***"
   ClientHeight    =   4905
   ClientLeft      =   2190
   ClientTop       =   2190
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5895
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
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   5655
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "MOVCAIXA.frx":0000
      Left            =   120
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Sistema\movim01.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modo de trabalho"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton Option6 
         Caption         =   "Mov. Lançam."
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Mov. Baixas"
         Height          =   315
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Movimentação"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Mov. cheques pré"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Cheques devolvidos"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Taxax dev. não pagas"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   960
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Data Final:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
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
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "MOVCAIXA.frx":0010
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "MOVCAIXA.frx":0020
      TabIndex        =   1
      Top             =   1920
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
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   5655
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Bindings        =   "MOVCAIXA.frx":1FC2
      Left            =   720
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Sistema\movim02.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin Crystal.CrystalReport CrystalReport3 
      Bindings        =   "MOVCAIXA.frx":1FD2
      Left            =   1320
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Sistema\movim03.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin Crystal.CrystalReport CrystalReport4 
      Bindings        =   "MOVCAIXA.frx":1FE2
      Left            =   1920
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin Crystal.CrystalReport CrystalReport5 
      Bindings        =   "MOVCAIXA.frx":1FF2
      Left            =   2520
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Sistema\movim05.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin Crystal.CrystalReport CrystalReport6 
      Bindings        =   "MOVCAIXA.frx":2002
      Left            =   3120
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Sistema\movim06.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "MOVCAIXA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim recRecordset1 As Recordset, recRecordset2 As Recordset
Set recRecordset1 = Data2.Recordset 'copy the recordset
teste2 = Date
teste3 = Date
If Option1.Value = True Then teste2 = CDate(MaskEdBox1)
If Option1.Value = True Then teste3 = CDate(MaskEdBox2)
If Option5.Value = True Then teste2 = CDate(MaskEdBox1)
If Option5.Value = True Then teste3 = CDate(MaskEdBox2)
If Option6.Value = True Then teste2 = CDate(MaskEdBox1)
If Option6.Value = True Then teste3 = CDate(MaskEdBox2)
teste2 = "cdate('" + Trim(Str(teste2)) + "')"
teste3 = "cdate('" + Trim(Str(teste3)) + "')"
If Option6.Value = True Then teste = "data >= " + teste2 + " and data <= " + teste3 + " and codf = 0 and ci = 0"
If Option5.Value = True Then teste = "PRE_DATA >= " + teste2 + " and PRE_DATA <= " + teste3 + " and codf <> 0 and ci <> 0"
If Option4.Value = True Then teste = "qtd_dev <> qtd_devp"
If Option3.Value = True Then teste = "qtd_dev <> 0"
If Option2.Value = True Then teste = "pre_data >= " + teste2 + " and codf <> 0 and ci <> 0"
If Option1.Value = True Then teste = "data >= " + teste2 + " and data <= " + teste3
recRecordset1.Filter = teste
Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type)
Set Data3.Recordset = recRecordset2 'assign back to original recordset object
DBGrid2.Visible = True
Command1.Visible = False
End Sub

Private Sub DBGrid1_Click()
Command1.Visible = True
DBGrid2.Visible = False

End Sub

Private Sub Command2_Click()
If Option1.Value = True Then CrystalReport1.PrintReport
If Option2.Value = True Then CrystalReport2.PrintReport
If Option3.Value = True Then CrystalReport3.PrintReport
If Option5.Value = True Then CrystalReport5.PrintReport
If Option6.Value = True Then CrystalReport6.PrintReport
End Sub

Private Sub Form_Activate()
MaskEdBox1 = Date
MaskEdBox2 = Date
End Sub

Private Sub MaskEdBox1_Change()
If MaskEdBox1.FormattedText = "" Then MaskEdBox1 = Date
Command1.Visible = True
DBGrid2.Visible = False

End Sub

Private Sub MaskEdBox2_Change()
If MaskEdBox2.FormattedText = "" Then MaskEdBox2 = Date
Command1.Visible = True
DBGrid2.Visible = False

End Sub

Private Sub Option1_Click()
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = True
MaskEdBox2.Enabled = True
End Sub

Private Sub Option2_Click()
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = False
MaskEdBox2.Enabled = False
End Sub

Private Sub Option3_Click()
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = False
MaskEdBox2.Enabled = False
End Sub

Private Sub Option4_Click()
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = False
MaskEdBox2.Enabled = False
End Sub

Private Sub Text3_Change()
Data1.Recordset.Seek ">=", Text3.Text
End Sub

Private Sub Option5_Click()
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = True
MaskEdBox2.Enabled = True

End Sub

Private Sub Option6_Click()
Command1.Visible = True
DBGrid2.Visible = False
MaskEdBox1.Enabled = True
MaskEdBox2.Enabled = True

End Sub
