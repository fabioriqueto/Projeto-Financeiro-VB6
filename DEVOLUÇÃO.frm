VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form DEVOLUÇÃO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerar devolução de documento                                                         ***Megatron***"
   ClientHeight    =   5655
   ClientLeft      =   2190
   ClientTop       =   615
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7470
   Begin VB.CommandButton Command7 
      Caption         =   "Retirar Pagto Taxa Dev."
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   29
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Acresc. Pagto Taxa Dev."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Confirmar Taxa de devolução de cheques"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   27
      Top             =   0
      Width           =   3255
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   255
      Left            =   2520
      TabIndex        =   26
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   327680
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Retirar Dev. Doc."
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   22
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Acrescentar Dev. Doc."
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      TabIndex        =   20
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Propriedades de Devolução"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   7215
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   6000
         TabIndex        =   33
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text6 
         DataField       =   "QTD_DEVP"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   31
         Top             =   960
         Width           =   375
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   "C:\Sistema\Bancod01.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CLIENTES"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "DEVOLUÇÃO.frx":0000
         DataField       =   "CI"
         DataSource      =   "Data2"
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   1320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   327680
         Enabled         =   0   'False
         ListField       =   "NOME"
         BoundColumn     =   "CODIGO"
         Text            =   "DBCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text5 
         DataField       =   "QTD_DEV"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   19
         Top             =   960
         Width           =   375
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         DataField       =   "PRE_DATA"
         DataSource      =   "Data2"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         DataField       =   "VALOR"
         DataSource      =   "Data2"
         Height          =   255
         Left            =   5040
         TabIndex        =   15
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text4 
         DataField       =   "CONTA"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         DataField       =   "CHEQUE"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         DataField       =   "BANCO"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "NOME"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Valor Taxas pagas:"
         Height          =   255
         Left            =   4560
         TabIndex        =   32
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Taxas pagas:"
         Height          =   255
         Left            =   5760
         TabIndex        =   30
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Cliente Responsável:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Qtd. Dev.:"
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Pré Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Valor:"
         Height          =   255
         Left            =   4440
         TabIndex        =   14
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Nº conta:"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Nº cheque:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Correntista:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Localiza documentos"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "Filtrar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
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
         RecordSource    =   "TROCA"
         Top             =   240
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
         Left            =   2040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "TROCA"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "DEVOLUÇÃO.frx":0010
         Height          =   1335
         Left            =   120
         OleObjectBlob   =   "DEVOLUÇÃO.frx":0020
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Label Label1 
         Caption         =   "Nº cheque:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Taxa de devolução de cheques:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "DEVOLUÇÃO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
Dim recRecordset1 As Recordset, recRecordset2 As Recordset
Set recRecordset1 = Data1.Recordset                        'copy the recordset
teste = "cheque = " + "'" + MaskEdBox1 + "'"
recRecordset1.Filter = teste
Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type) 'establish the filter
Set Data2.Recordset = recRecordset2                        'assign back to original recordset object
If Text5.Text <> "" Then
   DBGrid1.Visible = True
   Frame2.Visible = True
   Command3.Enabled = True
   Command4.Enabled = True
   Command6.Enabled = True
   Command7.Enabled = True
End If
End Sub

Private Sub Command2_Click()
DEVOLUÇÃO.Hide
End Sub

Private Sub Command3_Click()
Command1.Enabled = False
DBGrid1.Visible = False
Frame2.Visible = False
Command3.Enabled = False
Command4.Enabled = False
Text5.Text = Text5.Text + 1
If Text5.Text >= 2 Then Text5.Text = 2
Data1.Refresh
Data2.Refresh
End Sub

Private Sub Command4_Click()
If Text5 > Text6 Then
   Command1.Enabled = False
   DBGrid1.Visible = False
   Frame2.Visible = False
   Command3.Enabled = False
   Command4.Enabled = False
   Command6.Enabled = False
   Command7.Enabled = False
   Text5.Text = Text5.Text - 1
   If Text5.Text <= 0 Then Text5.Text = 0
   Data1.Refresh
   Data2.Refresh
Else
   MsgBox "Impossível retirar quantidade de devolução quando a quantidade de taxas pagas igual a quantidade de devolução", vbOKOnly, "Atenção!"
End If
End Sub

Private Sub Command5_Click()
Command5.Enabled = False
Frame1.Visible = True
DBGrid1.Visible = False
End Sub

Private Sub Command6_Click()
If Text5 > Text6 Then
   Command1.Enabled = False
   DBGrid1.Visible = False
   Frame2.Visible = False
   Command3.Enabled = False
   Command4.Enabled = False
   Command6.Enabled = False
   Command7.Enabled = False
   Text6.Text = Text6.Text + 1
   If Text6.Text >= 2 Then Text6.Text = 2
   Data1.Refresh
   Data2.Refresh
Else
   MsgBox "Impossível efetuar pagamento de taxa maior que a quantidade de devolução", vbOKOnly, "Atenção!"
End If
End Sub

Private Sub Command7_Click()
Command1.Enabled = False
DBGrid1.Visible = False
Frame2.Visible = False
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Text6.Text = Text6.Text - 1
If Text6.Text <= 0 Then Text6.Text = 0
Data1.Refresh
Data2.Refresh

End Sub

Private Sub Form_Load()
MaskEdBox4 = 6.5
End Sub

Private Sub MaskEdBox1_Change()
Command1.Enabled = True
DBGrid1.Visible = False
Frame2.Visible = False
Command3.Enabled = False
Command4.Enabled = False

End Sub

Private Sub MaskEdBox4_Change()
If MaskEdBox4.FormattedText = "" Then MaskEdBox4 = 0
Command5.Enabled = True
Frame1.Visible = False
Frame2.Visible = False
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Text1_Change()
MaskEdBox5 = MaskEdBox4 * Text6
End Sub
