VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Calculotrocal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculo para troca de cheque                                     *** Megatron ***"
   ClientHeight    =   6630
   ClientLeft      =   2505
   ClientTop       =   1005
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6240
   Begin VB.TextBox Text16b 
      DataField       =   "TAXATROCA"
      DataSource      =   "Data9"
      Height          =   285
      Left            =   6360
      TabIndex        =   60
      Text            =   "J Fin"
      Top             =   2160
      Width           =   735
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Financeiro"
      Top             =   1800
      Width           =   1140
   End
   Begin VB.TextBox Text17b 
      DataField       =   "JURO2"
      DataSource      =   "Data3"
      Height          =   285
      Left            =   6360
      TabIndex        =   59
      Text            =   "J Finan grava"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text18b 
      DataField       =   "LIQ2"
      DataSource      =   "Data3"
      Height          =   285
      Left            =   6360
      TabIndex        =   58
      Text            =   "Liq 2"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Data Data7 
      Caption         =   "Geral"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Data Data8 
      Caption         =   "Filtrado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Lançar"
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sair"
      Height          =   615
      Left            =   3240
      TabIndex        =   45
      Top             =   6000
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
      Left            =   0
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
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.TextBox Text6 
         DataField       =   "NOMEI"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   4440
         TabIndex        =   54
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text5 
         DataField       =   "FINANCEIRO"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   3720
         TabIndex        =   53
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Data Data6 
         Caption         =   "Data6"
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
         Top             =   2280
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text4 
         DataField       =   "TIPO"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   3960
         TabIndex        =   52
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   3720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   327680
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text3 
         DataField       =   "BANCO"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   6
         Top             =   3360
         Width           =   4695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Próxima fase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   4080
         Width           =   5775
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
         Left            =   2520
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
         TabIndex        =   3
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         DataField       =   "NOME"
         DataSource      =   "Data6"
         Height          =   285
         Left            =   1200
         TabIndex        =   44
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSMask.MaskEdBox MaskEdBox11 
         DataField       =   "CPF_CNPJ"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   1560
         TabIndex        =   1
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
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "CPF"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Caption         =   "Resultado pesquisa CPF / CNPJ"
         Height          =   1455
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   5775
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "Calculotrocal.frx":0000
            Height          =   1095
            Left            =   120
            OleObjectBlob   =   "Calculotrocal.frx":0010
            TabIndex        =   19
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
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "Calculotrocal.frx":0EDA
         DataField       =   "CODF"
         DataSource      =   "Data3"
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   3000
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   327680
         Enabled         =   0   'False
         ListField       =   "FINANCEIRO"
         BoundColumn     =   "CODIGO"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Calculotrocal.frx":0EEA
         DataField       =   "CI"
         DataSource      =   "Data3"
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   2640
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   327680
         Enabled         =   0   'False
         ListField       =   "NOME"
         BoundColumn     =   "CODIGO"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MaskEdBox15 
         Height          =   255
         Left            =   4200
         TabIndex        =   56
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "Calculotrocal.frx":0EFA
         Height          =   975
         Left            =   120
         OleObjectBlob   =   "Calculotrocal.frx":0F0A
         TabIndex        =   57
         Top             =   4680
         Width           =   5775
      End
      Begin VB.Label Label16 
         Caption         =   "Nome Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Nº Cheque:"
         Height          =   255
         Left            =   3120
         TabIndex        =   50
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Nº Conta:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Responsável:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Financeira:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Nome Titular:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2280
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Dados para cadastro"
      Height          =   4335
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Command6 
         Caption         =   "Fase anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   48
         Top             =   3720
         Width           =   5775
      End
      Begin VB.Frame Frame1 
         Caption         =   "Compensação"
         Height          =   1815
         Left            =   4560
         TabIndex        =   15
         Top             =   240
         Width           =   1335
         Begin VB.OptionButton Option1 
            Caption         =   "Isento"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Comp. D2"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Comp. D3"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Comp. D4"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   2160
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   327680
         Format          =   "0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   1800
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         DataField       =   "JURO"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   327680
         Format          =   "#,##0.000;(#,##0.000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         DataField       =   "PRE_DATA"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   1440
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   327680
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "DATA"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   327680
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         DataField       =   "VALOR"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   327680
         AutoTab         =   -1  'True
         OLEDragMode     =   1
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
         OLEDragMode     =   1
      End
      Begin VB.Frame Frame2 
         Caption         =   "Resultado Troca de Cheques"
         Height          =   1095
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   5775
         Begin MSMask.MaskEdBox MaskEdBox10 
            DataField       =   "LIQ"
            DataSource      =   "Data3"
            Height          =   375
            Left            =   3600
            TabIndex        =   24
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   327680
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox9 
            DataField       =   "DESC/ACRES"
            DataSource      =   "Data3"
            Height          =   255
            Left            =   4440
            TabIndex        =   25
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox8 
            DataField       =   "JUROS"
            DataSource      =   "Data3"
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox7 
            DataField       =   "CPMF"
            DataSource      =   "Data3"
            Height          =   255
            Left            =   2640
            TabIndex        =   27
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Caption         =   "Valor Liquido:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1440
            TabIndex        =   31
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "Desc. Total:"
            Height          =   255
            Left            =   3480
            TabIndex        =   30
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Juros:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "CPMF 0,30%:"
            Height          =   255
            Left            =   1560
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calcular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Taxa de Juros mensal:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Valor do documento:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "QTD. Dias:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Cheque pré p/:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Compensação:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1800
         Width           =   1095
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox14 
      DataField       =   "VALOR"
      DataSource      =   "Data8"
      Height          =   255
      Left            =   5280
      TabIndex        =   55
      Top             =   5520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   327680
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "Calculotrocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim juros1 As Integer
'Dim juros2 As Integer
'Dim doc1 As Integer
Dim dias1 As Integer
Dim dias2 As Integer
Dim semana1 As String
Dim semana2 As String

Private Sub Command1_Click()
semana1 = Left(MaskEdBox2.FormattedText, 7)
If semana1 = "Sábado," Then MaskEdBox2 = CDate(MaskEdBox2) + 2
If semana1 = "Domingo" Then MaskEdBox2 = CDate(MaskEdBox2) + 1
MaskEdBox6 = (CDate(MaskEdBox2) + 1) - CDate(MaskEdBox1)
MaskEdBox2 = (CDate(MaskEdBox1) - 1) + MaskEdBox6
If Option1 = True Then MaskEdBox5 = CDate(MaskEdBox2)
If Option2 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 2
If Option3 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 3
If Option4 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 4
semana2 = Left(MaskEdBox5.FormattedText, 7)
If Option2 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
ElseIf Option3 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Segunda" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
ElseIf Option4 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Segunda" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Terça-f" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
End If
dias2 = CDate(MaskEdBox5) - (CDate(MaskEdBox1) - 1)
juros3 = Val(MaskEdBox3.FormattedText) / 30
juros3 = juros3 * dias2
MaskEdBox7 = MaskEdBox4 * 0.003
MaskEdBox8 = MaskEdBox4 * (juros3 / 100)
MaskEdBox9 = (MaskEdBox4 * (juros3 / 100)) + (MaskEdBox4 * 0.003)
MaskEdBox10 = MaskEdBox4 - MaskEdBox9
Command1.Visible = False
Frame2.Visible = True
End Sub

Private Sub Command2_Click()
Dim recRecordset1s As Recordset, recRecordset2s As Recordset
Set recRecordset1s = Data7.Recordset 'copy the recordset
tts = MaskEdBox11
teste2s = Date
teste2s = "cdate('" + Trim(Str(teste2s)) + "')"
testes = "CPF_CNPJ = " + "'" + tts + "'" + " and pre_data >= " + teste2s
recRecordset1s.Filter = testes
Set recRecordset2s = recRecordset1s.OpenRecordset(recRecordset1s.Type)
Set Data8.Recordset = recRecordset2s 'assign back to original recordset object
Dim ttt1 As Integer
Dim ttt2 As Integer
If Data8.Recordset.EOF Then
Else
   Data8.Recordset.MoveFirst
   Do
      ttt2 = MaskEdBox14
      ttt1 = ttt1 + ttt2
      Data8.Recordset.MoveNext
      If Data8.Recordset.EOF = True Then Exit Do
   Loop
End If
MaskEdBox15 = ttt1
DBGrid2.Visible = True


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
Data3.Refresh
Calculotrocal.Hide
End Sub

Private Sub Command4_Click()
   Data9.Recordset.Seek "=", Val(DBCombo2.BoundText)
   Text17b = Text16b
   Text18b = MaskEdBox4 / ((100 - Text17b) / 100)

If Text2.Text <> "" And DBCombo1 <> "" And DBCombo2 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "           - " And MaskEdBox13.FormattedText <> "      " And Command1.Visible = False Then
   Data3.Recordset.Update
   Data1.Refresh
   Data3.Refresh
   Data3.Recordset.AddNew
   If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
   If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
   MaskEdBox1 = Date
   MaskEdBox2 = Date
   MaskEdBox3 = 0
   MaskEdBox4 = 0
   MaskEdBox6 = (CDate(MaskEdBox2) + 1) - CDate(MaskEdBox1)
   MaskEdBox7 = 0
   Command2.Enabled = False
   DBGrid1.Visible = False
   DBCombo1.Enabled = False
   DBCombo2.Enabled = False
   Frame5.Visible = False
   Frame3.Visible = True
   Text4.Text = "Troca de Cheques"
   MsgBox "Lançamento efetuado com sucesso!", vbOKOnly, "Atenção!"
Else
   MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
End If



End Sub

Private Sub Command5_Click()
If Text2.Text <> "" And DBCombo1 <> "" And DBCombo2 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "           - " And MaskEdBox13.FormattedText <> "      " Then
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

Private Sub DBCombo1_Click(Area As Integer)
Text6.Text = DBCombo1.Text
End Sub

Private Sub DBCombo2_Click(Area As Integer)
Text5.Text = DBCombo2.Text
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
Set db5 = OpenDatabase("c:\sistema\bancod01.mdb")
Set rs5 = db5.OpenRecordset("Financeiro")
Set Data9.Recordset = rs5
Data9.Recordset.Index = "primarykey"

Data3.Recordset.AddNew
If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
MaskEdBox1 = Date
MaskEdBox2 = Date
MaskEdBox3 = 0
MaskEdBox4 = 0
MaskEdBox6 = (CDate(MaskEdBox2) + 1) - CDate(MaskEdBox1)
MaskEdBox7 = 0
Text4.Text = "Troca de Cheques"
End Sub

Private Sub MaskEdBox1_Change()
If MaskEdBox1.FormattedText = "" Then MaskEdBox1 = Date
Command1.Visible = True
Frame2.Visible = False
End Sub

Private Sub MaskEdBox11_Change()
Command2.Enabled = True
DBGrid1.Visible = False
DBGrid2.Visible = False
DBCombo1.Enabled = False
DBCombo2.Enabled = False
End Sub

Private Sub MaskEdBox2_Change()
If MaskEdBox2.FormattedText = "" Then MaskEdBox2 = Date
Command1.Visible = True
Frame2.Visible = False

End Sub

Private Sub MaskEdBox3_Change()
If MaskEdBox3.FormattedText = "" Then MaskEdBox3 = 0
Command1.Visible = True
Frame2.Visible = False

End Sub

Private Sub MaskEdBox4_Change()
If MaskEdBox4.FormattedText = "" Then MaskEdBox4 = 0
Command1.Visible = True
Frame2.Visible = False

End Sub

Private Sub MaskEdBox6_Change()
If MaskEdBox6.FormattedText = "" Then MaskEdBox6 = 0
MaskEdBox2 = (CDate(MaskEdBox1) - 1) + MaskEdBox6
Command1.Visible = True
Frame2.Visible = False
End Sub

Private Sub Option1_Click()
semana1 = Left(MaskEdBox2.FormattedText, 7)
If semana1 = "Sábado," Then MaskEdBox2 = CDate(MaskEdBox2) + 2
If semana1 = "Domingo" Then MaskEdBox2 = CDate(MaskEdBox2) + 1
MaskEdBox6 = (CDate(MaskEdBox2) + 1) - CDate(MaskEdBox1)
MaskEdBox2 = (CDate(MaskEdBox1) - 1) + MaskEdBox6
If Option1 = True Then MaskEdBox5 = CDate(MaskEdBox2)
If Option2 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 2
If Option3 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 3
If Option4 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 4
semana2 = Left(MaskEdBox5.FormattedText, 7)
If Option2 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
ElseIf Option3 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Segunda" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
ElseIf Option4 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Segunda" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Terça-f" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
End If
dias2 = CDate(MaskEdBox5) - (CDate(MaskEdBox1) - 1)
juros3 = Val(MaskEdBox3.FormattedText) / 30
juros3 = juros3 * dias2
MaskEdBox7 = MaskEdBox4 * 0.003
MaskEdBox8 = MaskEdBox4 * (juros3 / 100)
MaskEdBox9 = (MaskEdBox4 * (juros3 / 100)) + (MaskEdBox4 * 0.003)
MaskEdBox10 = MaskEdBox4 - MaskEdBox9
Command1.Visible = True
Frame2.Visible = False

End Sub

Private Sub Option2_Click()
semana1 = Left(MaskEdBox2.FormattedText, 7)
If semana1 = "Sábado," Then MaskEdBox2 = CDate(MaskEdBox2) + 2
If semana1 = "Domingo" Then MaskEdBox2 = CDate(MaskEdBox2) + 1
MaskEdBox6 = (CDate(MaskEdBox2) + 1) - CDate(MaskEdBox1)
MaskEdBox2 = (CDate(MaskEdBox1) - 1) + MaskEdBox6
If Option1 = True Then MaskEdBox5 = CDate(MaskEdBox2)
If Option2 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 2
If Option3 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 3
If Option4 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 4
semana2 = Left(MaskEdBox5.FormattedText, 7)
If Option2 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
ElseIf Option3 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Segunda" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
ElseIf Option4 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Segunda" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Terça-f" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
End If
dias2 = CDate(MaskEdBox5) - (CDate(MaskEdBox1) - 1)
juros3 = Val(MaskEdBox3.FormattedText) / 30
juros3 = juros3 * dias2
MaskEdBox7 = MaskEdBox4 * 0.003
MaskEdBox8 = MaskEdBox4 * (juros3 / 100)
MaskEdBox9 = (MaskEdBox4 * (juros3 / 100)) + (MaskEdBox4 * 0.003)
MaskEdBox10 = MaskEdBox4 - MaskEdBox9
Command1.Visible = True
Frame2.Visible = False

End Sub

Private Sub Option3_Click()
semana1 = Left(MaskEdBox2.FormattedText, 7)
If semana1 = "Sábado," Then MaskEdBox2 = CDate(MaskEdBox2) + 2
If semana1 = "Domingo" Then MaskEdBox2 = CDate(MaskEdBox2) + 1
MaskEdBox6 = (CDate(MaskEdBox2) + 1) - CDate(MaskEdBox1)
MaskEdBox2 = (CDate(MaskEdBox1) - 1) + MaskEdBox6
If Option1 = True Then MaskEdBox5 = CDate(MaskEdBox2)
If Option2 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 2
If Option3 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 3
If Option4 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 4
semana2 = Left(MaskEdBox5.FormattedText, 7)
If Option2 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
ElseIf Option3 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Segunda" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
ElseIf Option4 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Segunda" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Terça-f" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
End If
dias2 = CDate(MaskEdBox5) - (CDate(MaskEdBox1) - 1)
juros3 = Val(MaskEdBox3.FormattedText) / 30
juros3 = juros3 * dias2
MaskEdBox7 = MaskEdBox4 * 0.003
MaskEdBox8 = MaskEdBox4 * (juros3 / 100)
MaskEdBox9 = (MaskEdBox4 * (juros3 / 100)) + (MaskEdBox4 * 0.003)
MaskEdBox10 = MaskEdBox4 - MaskEdBox9
Command1.Visible = True
Frame2.Visible = False

End Sub

Private Sub Option4_Click()
semana1 = Left(MaskEdBox2.FormattedText, 7)
If semana1 = "Sábado," Then MaskEdBox2 = CDate(MaskEdBox2) + 2
If semana1 = "Domingo" Then MaskEdBox2 = CDate(MaskEdBox2) + 1
MaskEdBox6 = (CDate(MaskEdBox2) + 1) - CDate(MaskEdBox1)
MaskEdBox2 = (CDate(MaskEdBox1) - 1) + MaskEdBox6
If Option1 = True Then MaskEdBox5 = CDate(MaskEdBox2)
If Option2 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 2
If Option3 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 3
If Option4 = True Then MaskEdBox5 = CDate(MaskEdBox2) + 4
semana2 = Left(MaskEdBox5.FormattedText, 7)
If Option2 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
ElseIf Option3 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Segunda" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
ElseIf Option4 = True Then
   If semana2 = "Domingo" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Sábado," Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Segunda" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
   If semana2 = "Terça-f" Then MaskEdBox5 = CDate(MaskEdBox5) + 2
End If
dias2 = CDate(MaskEdBox5) - (CDate(MaskEdBox1) - 1)
juros3 = Val(MaskEdBox3.FormattedText) / 30
juros3 = juros3 * dias2
MaskEdBox7 = MaskEdBox4 * 0.003
MaskEdBox8 = MaskEdBox4 * (juros3 / 100)
MaskEdBox9 = (MaskEdBox4 * (juros3 / 100)) + (MaskEdBox4 * 0.003)
MaskEdBox10 = MaskEdBox4 - MaskEdBox9
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

