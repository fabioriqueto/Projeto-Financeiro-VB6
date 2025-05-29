VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Calculosl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculo para Financiamento                                       *** Megatron ***"
   ClientHeight    =   7200
   ClientLeft      =   2505
   ClientTop       =   1005
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   6585
   Begin VB.TextBox Text18b 
      DataField       =   "LIQ2"
      DataSource      =   "Data3"
      Height          =   285
      Left            =   6840
      TabIndex        =   106
      Text            =   "Liq 2"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text17b 
      DataField       =   "JURO2"
      DataSource      =   "Data3"
      Height          =   285
      Left            =   6840
      TabIndex        =   105
      Text            =   "J Finan grava"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Financeiro"
      Top             =   960
      Width           =   1140
   End
   Begin VB.TextBox Text16b 
      DataField       =   "TAXATROCA"
      DataSource      =   "Data9"
      Height          =   285
      Left            =   6840
      TabIndex        =   104
      Text            =   "J Fin"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Data Data8 
      Caption         =   "Filtrado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Data Data7 
      Caption         =   "Geral"
      Connect         =   "Access"
      DatabaseName    =   "C:\Sistema\Bancod01.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados para consulta e cadastro"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox Text15 
         DataField       =   "CODF"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   2040
         TabIndex        =   100
         Text            =   "cod fin"
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text14 
         DataField       =   "CI"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   2040
         TabIndex        =   99
         Text            =   "cod resp"
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Fase Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   94
         Top             =   4920
         Width           =   6135
      End
      Begin VB.TextBox Text6 
         DataField       =   "NOMEI"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   4440
         TabIndex        =   47
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text5 
         DataField       =   "FINANCEIRO"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   3720
         TabIndex        =   46
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
         Left            =   2880
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
         TabIndex        =   45
         Text            =   "Tipo Transação"
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data Data4 
         Caption         =   "Data4"
         Connect         =   "Access"
         DatabaseName    =   "C:\Sistema\Bancod01.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3120
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
         TabIndex        =   19
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         DataField       =   "NOME"
         DataSource      =   "Data6"
         Height          =   285
         Left            =   1200
         TabIndex        =   41
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSMask.MaskEdBox MaskEdBox11 
         DataField       =   "CPF_CNPJ"
         DataSource      =   "Data3"
         Height          =   255
         Left            =   1560
         TabIndex        =   17
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
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "CPF"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Caption         =   "Resultado pesquisa CPF / CNPJ"
         Height          =   1455
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   6135
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "Calculosl.frx":0000
            Height          =   1095
            Left            =   120
            OleObjectBlob   =   "Calculosl.frx":0010
            TabIndex        =   37
            Top             =   240
            Visible         =   0   'False
            Width           =   5895
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Consultar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin MSDBCtls.DBCombo DBCombo2 
         Bindings        =   "Calculosl.frx":0EDA
         DataSource      =   "Data3"
         Height          =   315
         Left            =   1200
         TabIndex        =   21
         Top             =   3000
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   327680
         Enabled         =   0   'False
         ListField       =   "FINANCEIRO"
         BoundColumn     =   "CODIGO"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Calculosl.frx":0EEA
         DataSource      =   "Data3"
         Height          =   315
         Left            =   1200
         TabIndex        =   20
         Top             =   2640
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   327680
         Enabled         =   0   'False
         ListField       =   "NOME"
         BoundColumn     =   "CODIGO"
         Text            =   ""
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   1335
         Left            =   120
         TabIndex        =   35
         Top             =   3480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2355
         _Version        =   327680
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         Enabled         =   0   'False
         TabCaption(0)   =   "1ª Parcela"
         TabPicture(0)   =   "Calculosl.frx":0EFA
         Tab(0).ControlCount=   6
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Text3"
         Tab(0).Control(0).Enabled=   -1  'True
         Tab(0).Control(1)=   "MaskEdBox13"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "MaskEdBox12"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label15"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label14"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label16"
         Tab(0).Control(5).Enabled=   0   'False
         TabCaption(1)   =   "2ª Parcela"
         TabPicture(1)   =   "Calculosl.frx":0F16
         Tab(1).ControlCount=   6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Text7"
         Tab(1).Control(0).Enabled=   -1  'True
         Tab(1).Control(1)=   "MaskEdBox1"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "MaskEdBox2"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label3"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label2"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label1"
         Tab(1).Control(5).Enabled=   0   'False
         TabCaption(2)   =   "3ª Parcela"
         TabPicture(2)   =   "Calculosl.frx":0F32
         Tab(2).ControlCount=   6
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Text8"
         Tab(2).Control(0).Enabled=   -1  'True
         Tab(2).Control(1)=   "MaskEdBox3"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "MaskEdBox4"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Label6"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Label5"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Label4"
         Tab(2).Control(5).Enabled=   0   'False
         TabCaption(3)   =   "4ª Parcela"
         TabPicture(3)   =   "Calculosl.frx":0F4E
         Tab(3).ControlCount=   6
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "Label7"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Label8"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Label9"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "MaskEdBox6"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "MaskEdBox5"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "Text9"
         Tab(3).Control(5).Enabled=   0   'False
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   31
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   -73800
            MaxLength       =   50
            TabIndex        =   28
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   -73800
            MaxLength       =   50
            TabIndex        =   25
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox Text3 
            DataField       =   "BANCO"
            DataSource      =   "Data3"
            Height          =   285
            Left            =   -73800
            MaxLength       =   50
            TabIndex        =   22
            Top             =   480
            Width           =   4335
         End
         Begin MSMask.MaskEdBox MaskEdBox13 
            DataField       =   "CHEQUE"
            DataSource      =   "Data3"
            Height          =   255
            Left            =   -70920
            TabIndex        =   24
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
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
            Left            =   -74040
            TabIndex        =   23
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   327680
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   255
            Left            =   -70920
            TabIndex        =   27
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   327680
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   255
            Left            =   -74040
            TabIndex        =   26
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   327680
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   255
            Left            =   -70920
            TabIndex        =   30
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   327680
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   255
            Left            =   -74040
            TabIndex        =   29
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   327680
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   255
            Left            =   4080
            TabIndex        =   33
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
            _Version        =   327680
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox6 
            Height          =   255
            Left            =   960
            TabIndex        =   32
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   327680
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Caption         =   "Nº Cheque:"
            Height          =   255
            Left            =   3240
            TabIndex        =   59
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Nº Conta:"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Nome Banco:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Nº Cheque:"
            Height          =   255
            Left            =   -71760
            TabIndex        =   56
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Nº Conta:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   55
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Nome Banco:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   54
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Nº Cheque:"
            Height          =   255
            Left            =   -71760
            TabIndex        =   53
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Nº Conta:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   52
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Nome Banco:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   51
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Nº Cheque:"
            Height          =   255
            Left            =   -71760
            TabIndex        =   50
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Nº Conta:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   49
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "Nome Banco:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   48
            Top             =   480
            Width           =   975
         End
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   4560
         TabIndex        =   103
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
      Begin VB.Label Label12 
         Caption         =   "Responsável:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Financeira:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Nome Titular:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2280
         Width           =   975
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Lançar"
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   6600
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sair"
      Height          =   495
      Left            =   3600
      TabIndex        =   42
      Top             =   6600
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
      Top             =   4440
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
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TROCA"
      Top             =   720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados para calculo de financiamento"
      Height          =   6495
      Left            =   120
      TabIndex        =   60
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton Command1 
         Caption         =   "Próxima Fase"
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
         TabIndex        =   16
         Top             =   5880
         Width           =   6135
      End
      Begin VB.Frame Frame1a 
         Caption         =   "Resltado do Financiamento"
         Height          =   2055
         Left            =   120
         TabIndex        =   73
         Top             =   3720
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox Text10 
            DataField       =   "CPMF"
            DataSource      =   "Data3"
            Height          =   285
            Left            =   1560
            TabIndex        =   95
            Text            =   "cpmf / qtd par"
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSMask.MaskEdBox MaskEdBox17a 
            Height          =   375
            Left            =   3960
            TabIndex        =   74
            Top             =   120
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   327680
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox16a 
            Height          =   255
            Left            =   1320
            TabIndex        =   75
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
         Begin MSMask.MaskEdBox MaskEdBox15a 
            Height          =   255
            Left            =   4320
            TabIndex        =   76
            Top             =   1680
            Width           =   1695
            _ExtentX        =   2990
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
         Begin MSMask.MaskEdBox MaskEdBox14a 
            Height          =   255
            Left            =   4320
            TabIndex        =   77
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
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
         Begin MSMask.MaskEdBox MaskEdBox13a 
            Height          =   255
            Left            =   4320
            TabIndex        =   78
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
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
         Begin MSMask.MaskEdBox MaskEdBox12a 
            DataField       =   "VALOR"
            DataSource      =   "Data3"
            Height          =   255
            Left            =   4320
            TabIndex        =   79
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
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
         Begin MSMask.MaskEdBox MaskEdBox11a 
            Height          =   255
            Left            =   720
            TabIndex        =   80
            Top             =   1680
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            Enabled         =   0   'False
            Format          =   "dddddd"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox10a 
            Height          =   255
            Left            =   720
            TabIndex        =   81
            Top             =   1320
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            Enabled         =   0   'False
            Format          =   "dddddd"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox9a 
            Height          =   255
            Left            =   720
            TabIndex        =   82
            Top             =   960
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            Enabled         =   0   'False
            Format          =   "dddddd"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox8a 
            Height          =   255
            Left            =   720
            TabIndex        =   83
            Top             =   600
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            Enabled         =   0   'False
            Format          =   "dddddd"
            PromptChar      =   "_"
         End
         Begin VB.Label Label17a 
            Caption         =   "Valor Total:"
            Height          =   255
            Left            =   3000
            TabIndex        =   93
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label16a 
            Caption         =   "CPMF de 0.30:"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label15a 
            Caption         =   "Valor"
            Height          =   255
            Left            =   3840
            TabIndex        =   91
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label14a 
            Caption         =   "Valor"
            Height          =   255
            Left            =   3840
            TabIndex        =   90
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label13a 
            Caption         =   "Valor"
            Height          =   255
            Left            =   3840
            TabIndex        =   89
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label12a 
            Caption         =   "Valor"
            Height          =   255
            Left            =   3840
            TabIndex        =   88
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label11a 
            Caption         =   "Venc. 4"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label10a 
            Caption         =   "Venc. 3"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label9a 
            Caption         =   "Venc. 2"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label8a 
            Caption         =   "Venc. 1"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.CommandButton Command1a 
         Caption         =   "Calcular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   6135
      End
      Begin VB.Frame Frame4a 
         Caption         =   "Datas Vencimentos"
         Height          =   1815
         Left            =   1800
         TabIndex        =   67
         Top             =   1800
         Width           =   4455
         Begin VB.TextBox Text13 
            DataField       =   "JUROS"
            DataSource      =   "Data3"
            Height          =   285
            Left            =   1920
            TabIndex        =   98
            Text            =   "juros"
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox Text12 
            DataField       =   "DESC/ACRES"
            DataSource      =   "Data3"
            Height          =   285
            Left            =   1920
            TabIndex        =   97
            Text            =   "desc"
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton Command2a 
            Caption         =   "AUTO"
            Height          =   1335
            Left            =   4080
            TabIndex        =   68
            Top             =   360
            Width           =   255
         End
         Begin MSMask.MaskEdBox MaskEdBox2a 
            DataField       =   "PRE_DATA"
            DataSource      =   "Data3"
            Height          =   255
            Left            =   960
            TabIndex        =   11
            Top             =   360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            AllowPrompt     =   -1  'True
            Format          =   "dddddd"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox3a 
            Height          =   255
            Left            =   960
            TabIndex        =   12
            Top             =   720
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            Format          =   "dddddd"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox4a 
            Height          =   255
            Left            =   960
            TabIndex        =   13
            Top             =   1080
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            Format          =   "dddddd"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox5a 
            Height          =   255
            Left            =   960
            TabIndex        =   14
            Top             =   1440
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            Format          =   "dddddd"
            PromptChar      =   "_"
         End
         Begin VB.Label Label7a 
            Caption         =   "4 parcela"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label6a 
            Caption         =   "3 parcela"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label5a 
            Caption         =   "2 parcela"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label4a 
            Caption         =   "1 parcela"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame3a 
         Caption         =   "Tipo de taxa"
         Height          =   1455
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   1935
         Begin VB.OptionButton Option11a 
            Caption         =   "Taxa Postecipada"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton Option10a 
            Caption         =   "Taxa Antecipada"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option9a 
            Caption         =   "Taxa Mensal"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2a 
         Caption         =   "Forma Pagamento"
         Height          =   1815
         Left            =   120
         TabIndex        =   65
         Top             =   1800
         Width           =   1575
         Begin VB.OptionButton Option8a 
            Caption         =   "4 parcelas"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1440
            Width           =   1215
         End
         Begin VB.OptionButton Option7a 
            Caption         =   "3 parcelas"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton Option6a 
            Caption         =   "2 parcelas"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option5a 
            Caption         =   "1 parcela"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5a 
         Caption         =   "Dados Para Calculo"
         Height          =   1455
         Left            =   2160
         TabIndex        =   61
         Top             =   240
         Width           =   4095
         Begin VB.TextBox Text11 
            DataField       =   "LIQ"
            DataSource      =   "Data3"
            Height          =   285
            Left            =   1560
            TabIndex        =   96
            Text            =   "liq / parc"
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSMask.MaskEdBox MaskEdBox7a 
            Height          =   255
            Left            =   960
            TabIndex        =   3
            Top             =   1080
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox6a 
            DataField       =   "JURO"
            DataSource      =   "Data3"
            Height          =   255
            Left            =   960
            TabIndex        =   2
            Top             =   720
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            Format          =   "#,##0.000;(#,##0.000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1a 
            DataField       =   "DATA"
            DataSource      =   "Data3"
            Height          =   255
            Left            =   960
            TabIndex        =   1
            Top             =   360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   327680
            AllowPrompt     =   -1  'True
            Format          =   "dddddd"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1a 
            Caption         =   "Data Inicial:"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2a 
            Caption         =   "Juros:"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label3a 
            Caption         =   "Val. Finan.:"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1080
            Width           =   855
         End
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox7 
      DataField       =   "VALOR"
      DataSource      =   "Data8"
      Height          =   255
      Left            =   6840
      TabIndex        =   101
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   327680
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "Calculosl.frx":0F6A
      Height          =   975
      Left            =   6960
      OleObjectBlob   =   "Calculosl.frx":0F7A
      TabIndex        =   102
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "Calculosl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim juros1 As Integer
Dim dias As Integer
Private Sub Command1_Click()
Frame1.Visible = False
Frame3.Visible = True
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
      ttt2 = MaskEdBox7
      ttt1 = ttt1 + ttt2
      Data8.Recordset.MoveNext
      If Data8.Recordset.EOF = True Then Exit Do
   Loop
End If
MaskEdBox8 = ttt1


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
SSTab1.Enabled = True
End Sub

Private Sub Command3_Click()
Data3.Recordset.CancelUpdate
Calculosl.Hide
End Sub

Private Sub Command4_Click()
opicao = 0
If Option5a.Value = True Then
   If Text2.Text <> "" And Text14 <> "" And Text15 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "           - " And MaskEdBox13.FormattedText <> "      " And Command1a.Visible = False Then
      opicao = 1
   Else
      MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
   End If
ElseIf Option6a.Value = True Then
   If Text7.Text <> "" And Text2.Text <> "" And Text14 <> "" And Text15 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "           - " And MaskEdBox13.FormattedText <> "      " And Command1a.Visible = False And MaskEdBox2.FormattedText <> "           - " And MaskEdBox1.FormattedText <> "      " Then
      opicao = 1
   Else
      MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
   End If
ElseIf Option7a.Value = True Then
   If Text8.Text <> "" And Text7.Text <> "" And Text2.Text <> "" And Text14 <> "" And Text15 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "           - " And MaskEdBox13.FormattedText <> "      " And Command1a.Visible = False And MaskEdBox2.FormattedText <> "           - " And MaskEdBox1.FormattedText <> "      " And MaskEdBox4.FormattedText <> "           - " And MaskEdBox3.FormattedText <> "      " Then
      opicao = 1
   Else
      MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
   End If
ElseIf Option8a.Value = True Then
   If Text9.Text <> "" And Text8.Text <> "" And Text7.Text <> "" And Text2.Text <> "" And Text14 <> "" And Text15 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "           - " And MaskEdBox13.FormattedText <> "      " And Command1a.Visible = False And MaskEdBox2.FormattedText <> "           - " And MaskEdBox1.FormattedText <> "      " And MaskEdBox4.FormattedText <> "           - " And MaskEdBox3.FormattedText <> "      " And MaskEdBox6.FormattedText <> "           - " And MaskEdBox5.FormattedText <> "      " Then
      opicao = 1
   Else
      MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
   End If
End If




If opicao = 1 Then
   'fixos
       Data9.Recordset.Seek "=", Text15
       Text17b = Text16b
       Text18b = MaskEdBox12a / ((100 - Text17b) / 100)
       Datas = MaskEdBox1a
       juro = MaskEdBox6a
       liq = Text11
       Desc = Text12
       juros = Text13
       cpmf = Text10
       valor = MaskEdBox12a
       tipo = Text4
       cpf = MaskEdBox11
       nome = Text2
       codr = Text14
       nomer = Text6
       codf = Text15
       nomef = Text5
       bruto = MaskEdBox7a
If Option5a.Value = True Or Option6a.Value = True Or Option7a.Value = True Or Option8a.Value = True Then
   If Text2.Text <> "" And Text14 <> "" And Text15 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "           - " And MaskEdBox13.FormattedText <> "      " And Command1a.Visible = False Then
      Data3.Recordset.Update
      Data1.Refresh
      Data3.Refresh
      Data3.Recordset.AddNew
      MaskEdBox1a = Date
      MaskEdBox2a = Date
      MaskEdBox3a = Date
      MaskEdBox4a = Date
      MaskEdBox5a = Date
      MaskEdBox6a = 0
      MaskEdBox7a = 0
      If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
      If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
      Command2.Enabled = False
      DBGrid1.Visible = False
      DBCombo1.Enabled = False
      DBCombo2.Enabled = False
      Frame3.Visible = True
      Text4.Text = "Financiamento"
      MsgBox "Lançamento 1ª parcela efetuado com sucesso!", vbOKOnly, "Atenção!"
   Else
      MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
   End If
End If
If Option6a.Value = True Or Option7a.Value = True Or Option8a.Value = True Then
   'fixos
      MaskEdBox12a = valor
       MaskEdBox7a = bruto
       Text17b = Text16b
       Text18b = MaskEdBox12a / ((100 - Text17b) / 100)
      MaskEdBox1a = Datas
      MaskEdBox6a = juro
      Text11 = liq
      Text12 = Desc
      Text13 = juros
      Text10 = cpmf
      Text4 = tipo
      MaskEdBox11 = cpf
      Text2 = nome
      Text14 = Val(codr)
      Text6 = nomer
      Text15 = Val(codf)
      Text5 = nomef
    'variaveis
      MaskEdBox2a = MaskEdBox3a
      Text3 = Text7
      MaskEdBox12 = MaskEdBox2
      MaskEdBox13 = MaskEdBox1
   If Text2.Text <> "" And Text14 <> "" And Text15 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "           - " And MaskEdBox13.FormattedText <> "      " And Command1a.Visible = False And Text7.Text <> "" And MaskEdBox2.FormattedText <> "           - " And MaskEdBox1.FormattedText <> "      " Then
      Data3.Recordset.Update
      Data1.Refresh
      Data3.Refresh
      Data3.Recordset.AddNew
      MaskEdBox1a = Date
      MaskEdBox2a = Date
      MaskEdBox3a = Date
      MaskEdBox4a = Date
      MaskEdBox5a = Date
      MaskEdBox6a = 0
      MaskEdBox7a = 0
      If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
      If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
      Command2.Enabled = False
      DBGrid1.Visible = False
      DBCombo1.Enabled = False
      DBCombo2.Enabled = False
      Frame3.Visible = True
      Text4.Text = "Financiamento"
      MsgBox "Lançamento 2ª parcela efetuado com sucesso!", vbOKOnly, "Atenção!"
   Else
      MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
   End If
End If
If Option7a.Value = True Or Option8a.Value = True Then
   'fixos
      MaskEdBox12a = valor
       MaskEdBox7a = bruto
       Text17b = Text16b
       Text18b = MaskEdBox12a / ((100 - Text17b) / 100)
      MaskEdBox1a = Datas
      MaskEdBox6a = juro
      Text11 = liq
      Text12 = Desc
      Text13 = juros
      Text10 = cpmf
      Text4 = tipo
      MaskEdBox11 = cpf
      Text2 = nome
      Text14 = Val(codr)
      Text6 = nomer
      Text15 = Val(codf)
      Text5 = nomef
    'variaveis
      MaskEdBox2a = MaskEdBox4a
      Text3 = Text8
      MaskEdBox12 = MaskEdBox4
      MaskEdBox13 = MaskEdBox3
   If Text2.Text <> "" And Text14 <> "" And Text15 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "           - " And MaskEdBox13.FormattedText <> "      " And Command1a.Visible = False And Text7.Text <> "" And MaskEdBox2.FormattedText <> "           - " And MaskEdBox1.FormattedText <> "      " Then
      Data3.Recordset.Update
      Data1.Refresh
      Data3.Refresh
      Data3.Recordset.AddNew
      MaskEdBox1a = Date
      MaskEdBox2a = Date
      MaskEdBox3a = Date
      MaskEdBox4a = Date
      MaskEdBox5a = Date
      MaskEdBox6a = 0
      MaskEdBox7a = 0
      If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
      If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
      Command2.Enabled = False
      DBGrid1.Visible = False
      DBCombo1.Enabled = False
      DBCombo2.Enabled = False
      Frame3.Visible = True
      Text4.Text = "Financiamento"
      MsgBox "Lançamento 3ª parcela efetuado com sucesso!", vbOKOnly, "Atenção!"
   Else
      MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
   End If
End If
If Option8a.Value = True Then
   'fixos
      MaskEdBox12a = valor
       MaskEdBox7a = bruto
       Text17b = Text16b
       Text18b = MaskEdBox12a / ((100 - Text17b) / 100)
      MaskEdBox1a = Datas
      MaskEdBox6a = juro
      Text11 = liq
      Text12 = Desc
      Text13 = juros
      Text10 = cpmf
      Text4 = tipo
      MaskEdBox11 = cpf
      Text2 = nome
      Text14 = Val(codr)
      Text6 = nomer
      Text15 = Val(codf)
      Text5 = nomef
    'variaveis
      MaskEdBox2a = MaskEdBox5a
      Text3 = Text9
      MaskEdBox12 = MaskEdBox6
      MaskEdBox13 = MaskEdBox5
   If Text2.Text <> "" And Text14 <> "" And Text15 <> "" And Text3.Text <> "" And MaskEdBox12.FormattedText <> "           - " And MaskEdBox13.FormattedText <> "      " And Command1a.Visible = False And Text7.Text <> "" And MaskEdBox2.FormattedText <> "           - " And MaskEdBox1.FormattedText <> "      " Then
      Data3.Recordset.Update
      Data1.Refresh
      Data3.Refresh
      Data3.Recordset.AddNew
      MaskEdBox1a = Date
      MaskEdBox2a = Date
      MaskEdBox3a = Date
      MaskEdBox4a = Date
      MaskEdBox5a = Date
      MaskEdBox6a = 0
      MaskEdBox7a = 0
      If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
      If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
      Command2.Enabled = False
      DBGrid1.Visible = False
      DBCombo1.Enabled = False
      DBCombo2.Enabled = False
      Frame3.Visible = True
      Text4.Text = "Financiamento"
      MsgBox "Lançamento 4ª parcela efetuado com sucesso!", vbOKOnly, "Atenção!"
   Else
      MsgBox "Preencha todos os campos", vbOKOnly, "Atenção!"
   End If
End If
Frame3.Visible = False
Frame1.Visible = True
End If
End Sub

Private Sub Command5_Click()
Frame3.Visible = False
Frame1.Visible = True
End Sub

Private Sub DBCombo1_Click(Area As Integer)
Text6.Text = DBCombo1.Text
Text14.Text = DBCombo1.BoundText
End Sub

Private Sub DBCombo2_Click(Area As Integer)
Text5.Text = DBCombo2.Text
Text15.Text = DBCombo2.BoundText
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
Text4.Text = "Financiamento"
MaskEdBox1a = Date
MaskEdBox2a = Date
MaskEdBox3a = Date
MaskEdBox4a = Date
MaskEdBox5a = Date
MaskEdBox6a = 0
MaskEdBox7a = 0
End Sub

Private Sub MaskEdBox11_Change()
Command2.Enabled = True
DBGrid1.Visible = False
DBCombo1.Enabled = False
DBCombo2.Enabled = False
SSTab1.Enabled = False
End Sub

Private Sub MaskEdBox16a_Change()
If Option5a.Value = True Then Text10 = MaskEdBox16a
If Option6a.Value = True Then Text10 = MaskEdBox16a / 2
If Option7a.Value = True Then Text10 = MaskEdBox16a / 3
If Option8a.Value = True Then Text10 = MaskEdBox16a / 4
If Option5a.Value = True Then Text11 = MaskEdBox7a
If Option6a.Value = True Then Text11 = MaskEdBox7a / 2
If Option7a.Value = True Then Text11 = MaskEdBox7a / 3
If Option8a.Value = True Then Text11 = MaskEdBox7a / 4



End Sub

Private Sub Option5_Click()
If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
DBCombo1.Enabled = False
DBCombo2.Enabled = False
Command2.Enabled = False
SSTab1.Enabled = False
End Sub

Private Sub Option6_Click()
If Option5.Value = True Then MaskEdBox11.Mask = "###.###.###-##"
If Option6.Value = True Then MaskEdBox11.Mask = "###.###.###/####-##"
DBCombo1.Enabled = False
DBCombo2.Enabled = False
Command2.Enabled = False
SSTab1.Enabled = False
End Sub

Private Sub Command1a_Click()
If Option5a.Value = True Then
   dias = (CDate(MaskEdBox2a) + 1) - CDate(MaskEdBox1a)
End If
If Option6a.Value = True Then
   dias = (CDate(MaskEdBox3a) + 1) - CDate(MaskEdBox1a)
End If
If Option7a.Value = True Then
   dias = (CDate(MaskEdBox4a) + 1) - CDate(MaskEdBox1a)
End If
If Option8a.Value = True Then
   dias = (CDate(MaskEdBox5a) + 1) - CDate(MaskEdBox1a)
End If
MaskEdBox8a = MaskEdBox2a
MaskEdBox9a = MaskEdBox3a
MaskEdBox10a = MaskEdBox4a
MaskEdBox11a = MaskEdBox5a
Command1a.Visible = False
If Option5a.Value = True Then
   If Option10a.Value = True Then
      juros1 = (MaskEdBox6a / 30) * dias
      MaskEdBox17a = (MaskEdBox7a / ((100 - juros1) / 100)) / 0.997
      MaskEdBox16a = MaskEdBox17a * 0.003
   ElseIf Option11a.Value = True Then
      juros1 = (MaskEdBox6a / 30) * dias
      MaskEdBox17a = (MaskEdBox7a * ((juros1 / 100) + 1)) / 0.997
      MaskEdBox16a = MaskEdBox17a * 0.003
   End If
   MaskEdBox12a = MaskEdBox17a
   Label8a.Visible = True
   Label12a.Visible = True
   MaskEdBox8a.Visible = True
   MaskEdBox12a.Visible = True
   Label9a.Visible = False
   Label10a.Visible = False
   Label11a.Visible = False
   Label13a.Visible = False
   Label14a.Visible = False
   Label15a.Visible = False
   MaskEdBox9a.Visible = False
   MaskEdBox10a.Visible = False
   MaskEdBox11a.Visible = False
   MaskEdBox13a.Visible = False
   MaskEdBox14a.Visible = False
   MaskEdBox15a.Visible = False
Else
   If Option6a.Value = True Then
      dias = (CDate(MaskEdBox3a) + 1) - CDate(MaskEdBox1a)
      juros1 = (MaskEdBox6a / 30) * dias
      MaskEdBox17a = (MaskEdBox7a * ((juros1 / 100) + 1)) / 0.997
      MaskEdBox16a = MaskEdBox17a * 0.003
      MaskEdBox12a = MaskEdBox17a / 2
      MaskEdBox13a = MaskEdBox17a / 2
      Label8a.Visible = True
      Label12a.Visible = True
      MaskEdBox8a.Visible = True
      MaskEdBox12a.Visible = True
      Label9a.Visible = True
      Label10a.Visible = False
      Label11a.Visible = False
      Label13a.Visible = True
      Label14a.Visible = False
      Label15a.Visible = False
      MaskEdBox9a.Visible = True
      MaskEdBox10a.Visible = False
      MaskEdBox11a.Visible = False
      MaskEdBox13a.Visible = True
      MaskEdBox14a.Visible = False
      MaskEdBox15a.Visible = False
   End If
   If Option7a.Value = True Then
      dias = (CDate(MaskEdBox4a) + 1) - CDate(MaskEdBox1a)
      juros1 = (MaskEdBox6a / 30) * dias
      MaskEdBox17a = (MaskEdBox7a * ((juros1 / 100) + 1)) / 0.997
      MaskEdBox16a = MaskEdBox17a * 0.003
      MaskEdBox12a = MaskEdBox17a / 3
      MaskEdBox13a = MaskEdBox17a / 3
      MaskEdBox14a = MaskEdBox17a / 3
      Label8a.Visible = True
      Label12a.Visible = True
      MaskEdBox8a.Visible = True
      MaskEdBox12a.Visible = True
      Label9a.Visible = True
      Label10a.Visible = True
      Label11a.Visible = False
      Label13a.Visible = True
      Label14a.Visible = True
      Label15a.Visible = False
      MaskEdBox9a.Visible = True
      MaskEdBox10a.Visible = True
      MaskEdBox11a.Visible = False
      MaskEdBox13a.Visible = True
      MaskEdBox14a.Visible = True
      MaskEdBox15a.Visible = False
   End If
   If Option8a.Value = True Then
      dias = (CDate(MaskEdBox5a) + 1) - CDate(MaskEdBox1a)
      juros1 = (MaskEdBox6a / 30) * dias
      MaskEdBox17a = (MaskEdBox7a * ((juros1 / 100) + 1)) / 0.997
      MaskEdBox16a = MaskEdBox17a * 0.003
      MaskEdBox12a = MaskEdBox17a / 4
      MaskEdBox13a = MaskEdBox17a / 4
      MaskEdBox14a = MaskEdBox17a / 4
      MaskEdBox15a = MaskEdBox17a / 4
      Label8a.Visible = True
      Label12a.Visible = True
      MaskEdBox8a.Visible = True
      MaskEdBox12a.Visible = True
      Label9a.Visible = True
      Label10a.Visible = True
      Label11a.Visible = True
      Label13a.Visible = True
      Label14a.Visible = True
      Label15a.Visible = True
      MaskEdBox9a.Visible = True
      MaskEdBox10a.Visible = True
      MaskEdBox11a.Visible = True
      MaskEdBox13a.Visible = True
      MaskEdBox14a.Visible = True
      MaskEdBox15a.Visible = True
   End If
End If
If Option5a.Value = True Then Text10 = MaskEdBox16a
If Option6a.Value = True Then Text10 = MaskEdBox16a / 2
If Option7a.Value = True Then Text10 = MaskEdBox16a / 3
If Option8a.Value = True Then Text10 = MaskEdBox16a / 4
If Option5a.Value = True Then Text11 = MaskEdBox7a
If Option6a.Value = True Then Text11 = MaskEdBox7a / 2
If Option7a.Value = True Then Text11 = MaskEdBox7a / 3
If Option8a.Value = True Then Text11 = MaskEdBox7a / 4
Text12 = MaskEdBox12a - Val(Text11)
Text13 = Text12 - Text10


Frame1a.Visible = True

End Sub

Private Sub Command2a_Click()
wdia1 = Day(CDate(MaskEdBox2a))
wmes1 = Month(CDate(MaskEdBox2a))
wano1 = Year(CDate(MaskEdBox2a))
wmes1 = Val(wmes1) + 1
If Val(wmes1) > 12 Then
   wmes1 = 1
   wano1 = Val(wano1) + 1
End If
MaskEdBox3a = Str(wdia1) + "-" + Str(wmes1) + "-" + Str(wano1)
wteste = MaskEdBox3a.FormattedText
If Left(Trim(wteste), 5) = "31- 2" Then MaskEdBox3a = "28-2-" + Str(wano1)
If Left(Trim(wteste), 5) = "30- 2" Then MaskEdBox3a = "28-2-" + Str(wano1)
If Left(Trim(wteste), 5) = "29- 2" Then MaskEdBox3a = "28-2-" + Str(wano1)

wdia2 = Day(CDate(MaskEdBox2a))
wmes2 = Month(CDate(MaskEdBox3a))
wano2 = Year(CDate(MaskEdBox3a))
wmes2 = Val(wmes2) + 1
If Val(wmes2) > 12 Then
   wmes2 = 1
   wano2 = Val(wano2) + 1
End If
MaskEdBox4a = Str(wdia2) + "-" + Str(wmes2) + "-" + Str(wano2)
wteste = MaskEdBox4a.FormattedText
If Left(Trim(wteste), 5) = "31- 2" Then MaskEdBox4a = "28-2-" + Str(wano2)
If Left(Trim(wteste), 5) = "30- 2" Then MaskEdBox4a = "28-2-" + Str(wano2)
If Left(Trim(wteste), 5) = "29- 2" Then MaskEdBox4a = "28-2-" + Str(wano2)

wdia3 = Day(CDate(MaskEdBox2a))
wmes3 = Month(CDate(MaskEdBox4a))
wano3 = Year(CDate(MaskEdBox4a))
wmes3 = Val(wmes3) + 1
If Val(wmes3) > 12 Then
   wmes3 = 1
   wano3 = Val(wano3) + 1
End If
MaskEdBox5a = Str(wdia3) + "-" + Str(wmes3) + "-" + Str(wano3)
wteste = MaskEdBox5a.FormattedText
If Left(Trim(wteste), 5) = "31- 2" Then MaskEdBox5a = "1-3-" + Str(wano3)
If Left(Trim(wteste), 5) = "30- 2" Then MaskEdBox5a = "1-3-" + Str(wano3)
If Left(Trim(wteste), 5) = "29- 2" Then MaskEdBox5a = "1-3-" + Str(wano3)



End Sub

Private Sub MaskEdBox1a_Change()
If MaskEdBox1a.FormattedText = "" Then MaskEdBox1a = Date
Command1a.Visible = True
Frame1a.Visible = False

End Sub

Private Sub MaskEdBox2a_Change()
If MaskEdBox2a.FormattedText = "" Then MaskEdBox2a = Date
Command1a.Visible = True
Frame1a.Visible = False

End Sub

Private Sub MaskEdBox3a_Change()
If MaskEdBox3a.FormattedText = "" Then MaskEdBox3a = Date
Command1a.Visible = True
Frame1a.Visible = False

End Sub

Private Sub MaskEdBox4a_Change()
If MaskEdBox4a.FormattedText = "" Then MaskEdBox4a = Date
Command1a.Visible = True
Frame1a.Visible = False

End Sub

Private Sub MaskEdBox5a_Change()
If MaskEdBox5a.FormattedText = "" Then MaskEdBox5a = Date
Command1a.Visible = True
Frame1a.Visible = False

End Sub

Private Sub MaskEdBox6a_Change()
Command1a.Visible = True
Frame1a.Visible = False
If MaskEdBox6a.FormattedText = "" Then MaskEdBox6a = 0
End Sub

Private Sub MaskEdBox7a_Change()
Command1a.Visible = True
Frame1a.Visible = False
If MaskEdBox7a.FormattedText = "" Then MaskEdBox7a = 0
If Option5a.Value = True Then Text11 = MaskEdBox7a
If Option6a.Value = True Then Text11 = MaskEdBox7a / 2
If Option7a.Value = True Then Text11 = MaskEdBox7a / 3
If Option8a.Value = True Then Text11 = MaskEdBox7a / 4

End Sub

Private Sub Option10a_Click()
Command1a.Visible = True
Frame1a.Visible = False

End Sub

Private Sub Option11a_Click()
Command1a.Visible = True
Frame1a.Visible = False

End Sub

Private Sub Option5a_Click()
Option9a.Enabled = False
Option10a.Enabled = True
Option11a.Enabled = True
Option10a.Value = True
MaskEdBox2a.Enabled = True
MaskEdBox3a.Enabled = False
MaskEdBox4a.Enabled = False
MaskEdBox5a.Enabled = False
Command1a.Visible = True
Frame1a.Visible = False
End Sub

Private Sub Option6a_Click()
Option9a.Enabled = True
Option10a.Enabled = False
Option11a.Enabled = False
Option9a.Value = True
MaskEdBox2a.Enabled = True
MaskEdBox3a.Enabled = True
MaskEdBox4a.Enabled = False
MaskEdBox5a.Enabled = False
Command1a.Visible = True
Frame1a.Visible = False
End Sub

Private Sub Option7a_Click()
Option9a.Enabled = True
Option10a.Enabled = False
Option11a.Enabled = False
Option9a.Value = True
MaskEdBox2a.Enabled = True
MaskEdBox3a.Enabled = True
MaskEdBox4a.Enabled = True
MaskEdBox5a.Enabled = False
Command1a.Visible = True
Frame1a.Visible = False
End Sub

Private Sub Option8a_Click()
Option9a.Enabled = True
Option10a.Enabled = False
Option11a.Enabled = False
Option9a.Value = True
MaskEdBox2a.Enabled = True
MaskEdBox3a.Enabled = True
MaskEdBox4a.Enabled = True
MaskEdBox5a.Enabled = True
Command1a.Visible = True
Frame1a.Visible = False
End Sub

Private Sub Option9a_Click()
Command1a.Visible = True
Frame1a.Visible = False

End Sub

