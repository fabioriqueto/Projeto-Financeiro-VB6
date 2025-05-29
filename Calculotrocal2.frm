VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Calculotrocal2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculo para troca de cheque                                     *** Megatron ***"
   ClientHeight    =   5700
   ClientLeft      =   2505
   ClientTop       =   1005
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6255
   Begin VB.CommandButton Command4 
      Caption         =   "Lançar"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sair"
      Height          =   615
      Left            =   3240
      TabIndex        =   10
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados para consulta e cadastro"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   4080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Data Data6 
         Caption         =   "Data6"
         Connect         =   "Access"
         DatabaseName    =   "C:\Sistema\Bancod01.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TROCA"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text4 
         DataField       =   "TIPO"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   3960
         TabIndex        =   22
         Text            =   "Tipo Transação"
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         DataField       =   "CHEQUE"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   3720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox12 
         DataField       =   "CONTA"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   3720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "########-#"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text3 
         DataField       =   "BANCO"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   19
         Top             =   3360
         Width           =   4695
      End
      Begin VB.Data Data4 
         Caption         =   "Data4"
         Connect         =   "Access"
         DatabaseName    =   "C:\Sistema\Bancod01.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CLIENTES"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data5 
         Caption         =   "Data5"
         Connect         =   "Access"
         DatabaseName    =   "C:\Sistema\Bancod01.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "FINANCEIRO"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         DataField       =   "NOME"
         DataSource      =   "Data3"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2280
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         DataField       =   "NOME"
         DataSource      =   "Data6"
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSMask.MaskEdBox MaskEdBox11 
         DataField       =   "CPF_CNPJ"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   327680
         PromptChar      =   "_"
      End
      Begin VB.OptionButton Option6 
         Caption         =   "CNPJ"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "CPF"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Caption         =   "Resultado pesquisa CPF / CNPJ"
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   5775
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "Calculotrocal2.frx":0000
            Height          =   1095
            Left            =   120
            OleObjectBlob   =   "Calculotrocal2.frx":0010
            TabIndex        =   3
            Top             =   240
            Visible         =   0   'False
            Width           =   5535
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Consultar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "Calculotrocal2.frx":0EDA
         DataField       =   "CODF"
         DataSource      =   "Data3"
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   3000
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   327680
         Enabled         =   0   'False
         ListField       =   "FINANCEIRO"
         BoundColumn     =   "CODIGO"
         Text            =   "DBCombo2"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Calculotrocal2.frx":0EEA
         DataField       =   "CI"
         DataSource      =   "Data3"
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   2640
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   327680
         Enabled         =   0   'False
         ListField       =   "NOME"
         BoundColumn     =   "CODIGO"
         Text            =   "DBCombo1"
      End
      Begin VB.Label Label16 
         Caption         =   "Nome Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Nº Cheque:"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Nº Conta:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Responsável:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Financeira:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Nome Titular:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
   End
End
Attribute VB_Name = "Calculotrocal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Dim teste As String
DBGrid1.Visible = True
Command2.Enabled = False


Dim recRecordset1 As Recordset, recRecordset2 As Recordset
Set recRecordset1 = Data1.Recordset                        'copy the recordset
teste = "cpf_cnpj = " + "'" + MaskEdBox11 + "'" + " and qtd_dev <> 0"
recRecordset1.Filter = teste
Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type) 'establish the filter
Set Data2.Recordset = recRecordset2                        'assign back to original recordset object


Dim recRecordset6 As Recordset, recRecordset7 As Recordset
Set recRecordset6 = Data1.Recordset                        'copy the recordset
teste = "cpf_cnpj = " + "'" + MaskEdBox11 + "'"
recRecordset1.Filter = teste
Set recRecordset7 = recRecordset6.OpenRecordset(recRecordset6.Type) 'establish the filter
Set Data6.Recordset = recRecordset7                        'assign back to original recordset object


Text2.Text = Text1.Text
If Text1.Text = "" Then Text2.Enabled = True
If Text1.Text <> "" Then Text2.Enabled = False
'Data2.RecordSource = "select [cod],[BANCO],[CONTA],[CHEQUE],[VALOR],[NOME],[TIPO],[DEVOLU],[DATA],[PARA] from [DEVOLUÇÃO] where [CPF_CNPJ]='" & Text1.Text & "'"
'Data2.Refresh
DBCombo1.Enabled = True
DBCombo2.Enabled = True
End Sub

Private Sub Command3_Click()
Data3.Recordset.CancelUpdate
Calculotrocal2.Hide
End Sub

Private Sub Command4_Click()
If Text2.Text <> "" And DBCombo1 <> "" And DBCombo2 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "        - " And MaskEdBox13.FormattedText <> "      " Then
   Data3.Recordset.Update
   Data1.Refresh
   Data3.Refresh
   Data3.Recordset.MoveFirst
   Data3.Recordset.AddNew
   If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
   If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
   Command2.Enabled = False
   DBGrid1.Visible = False
   DBCombo1.Enabled = False
   DBCombo2.Enabled = False
   Text4.Text = "Financiamento"
   Text5.Text = Date
   MsgBox "Lançamento efetuado com sucesso!", vbOKOnly, "Atenção!"
Else
   MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
End If
End Sub

Private Sub Command5_Click()
If Text2.Text <> "" And DBCombo1 <> "" And DBCombo2 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "        - " And MaskEdBox13.FormattedText <> "      " Then
   Frame3.Visible = False
   Frame5.Visible = True
Else
   MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
End If

End Sub

Private Sub Command6_Click()
Frame5.Visible = False
Frame3.Visible = True
End Sub

Private Sub Form_Activate()
'indexa banco de dados de clientes
Set db2 = OpenDatabase("c:\sistema\bancod01.mdb")
Set rs2 = db2.OpenRecordset("clientes")
Set Data4.Recordset = rs2
Data4.Recordset.Index = "inome"
'indexa banco de dados de financeiras
Set db3 = OpenDatabase("c:\sistema\bancod01.mdb")
Set rs3 = db3.OpenRecordset("financeiro")
Set Data5.Recordset = rs3
Data5.Recordset.Index = "ifinanceiro"
Data3.Recordset.AddNew
If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
Text4.Text = "Financiamento"
Text5.Text = Date
End Sub

Private Sub MaskEdBox1_Change()
Command1.Visible = True
Frame2.Visible = False
End Sub

Private Sub MaskEdBox11_Change()
Command2.Enabled = True
DBGrid1.Visible = False
DBCombo1.Enabled = False
DBCombo2.Enabled = False
End Sub

Private Sub MaskEdBox2_Change()
Command1.Visible = True
Frame2.Visible = False

End Sub

Private Sub MaskEdBox3_Change()
Command1.Visible = True
Frame2.Visible = False

End Sub

Private Sub MaskEdBox4_Change()
Command1.Visible = True
Frame2.Visible = False

End Sub

Private Sub Option5_Click()
If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
DBCombo1.Enabled = False
DBCombo2.Enabled = False
Command2.Enabled = False
End Sub

Private Sub Option6_Click()
If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
DBCombo1.Enabled = False
DBCombo2.Enabled = False
Command2.Enabled = False
End Sub
