VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Acumulo 
   Caption         =   "Resultado de acumulo da pesquisa por CPF"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox MaskEdBox2 
      DataField       =   "VALOR"
      DataSource      =   "Data3"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   327680
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      _Version        =   327680
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Data Data2 
      Caption         =   "Geral"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   720
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Data Data3 
      Caption         =   "Filtrado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continuar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   2040
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "Acumulo.frx":0000
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "Acumulo.frx":0010
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Acumulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Acumulo.Hide
End Sub

Private Sub Form_Activate()
Dim recRecordset1 As Recordset, recRecordset2 As Recordset
Set recRecordset1 = Data2.Recordset 'copy the recordset
'tt = "111.111.111-11"
teste2 = Date
teste2 = "cdate('" + Trim(Str(teste2)) + "')"
teste = "CPF_CNPJ = " + "'" + tt + "'" + " and pre_data >= " + teste2
recRecordset1.Filter = teste
Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type)
Set Data3.Recordset = recRecordset2 'assign back to original recordset object
Dim ttt1 As Integer
Dim ttt2 As Integer
If Data3.Recordset.EOF Then
   Acumulo.Hide
Else
   Data3.Recordset.MoveFirst
   Do
      ttt2 = MaskEdBox2
      ttt1 = ttt1 + ttt2
      Data3.Recordset.MoveNext
      If Data3.Recordset.EOF = True Then Exit Do
   Loop
End If
MaskEdBox1 = ttt1
End Sub
