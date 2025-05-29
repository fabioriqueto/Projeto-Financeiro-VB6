VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Lançamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamentos                                    ***Megatron***"
   ClientHeight    =   1815
   ClientLeft      =   2190
   ClientTop       =   3180
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4680
   Begin VB.TextBox Text5 
      DataField       =   "VALOR"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   2400
      Width           =   1140
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      DataField       =   "VALOR"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   327680
      Format          =   "R$#,##0.00;(R$#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text4 
      DataField       =   "TIPO"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      DataField       =   "PRE_DATA"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      DataField       =   "DATA"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sair"
      Height          =   495
      Left            =   3600
      MaskColor       =   &H8000000F&
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Efetuar Lançamento"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Limpar Campos"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "LIQ"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   327680
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      Format          =   "R$#,##0.00;(R$#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text1 
      DataField       =   "NOME"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   480
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Lançamento de &Débito"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Lançamento de &Crédito"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Valor:"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Referente à:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Lançamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MaskEdBox1 = 0
Text1.Text = ""
End Sub

Private Sub Command2_Click()
MaskEdBox2 = MaskEdBox1
If Text1.Text <> "" And MaskEdBox1 <> 0 Then
   Data1.Recordset.Update
   Data1.Refresh
   Data1.Recordset.AddNew
   MaskEdBox1 = 0
   Text1.Text = ""
   Text2.Text = Date
   Text3.Text = Date
   If Option1.Value = True Then Text4.Text = "Lançamento de Crédito"
   If Option2.Value = True Then Text4.Text = "Lançamento de Débito"
   MsgBox "Lançamento efetuado com sucesso!", vbOKOnly, "Atenção!"
Else
   MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
End If
End Sub

Private Sub Command3_Click()
Data1.Recordset.CancelUpdate
Lançamentos.Hide
End Sub

Private Sub Form_Activate()
Data1.Recordset.AddNew
MaskEdBox1 = 0
Text2.Text = Date
Text3.Text = Date
If Option1.Value = True Then Text4.Text = "Lançamento de Crédito"
If Option2.Value = True Then Text4.Text = "Lançamento de Débito"

End Sub

Private Sub MaskEdBox1_Change()
If MaskEdBox1.FormattedText = "" Then MaskEdBox1 = 0
If Option1.Value = True And Val(MaskEdBox1) < 0 Then MaskEdBox1 = MaskEdBox1 * (-1)
If Option2.Value = True And Val(MaskEdBox1) > 0 Then MaskEdBox1 = MaskEdBox1 * (-1)
Text5 = MaskEdBox1
End Sub

Private Sub Option1_Click()
Text4.Text = "Lançamento de Crédito"

End Sub

Private Sub Option2_Click()
Text4.Text = "Lançamento de Débito"

End Sub
