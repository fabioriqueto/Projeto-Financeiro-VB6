VERSION 5.00
Begin VB.Form Principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brusk Factoring"
   ClientHeight    =   5895
   ClientLeft      =   195
   ClientTop       =   480
   ClientWidth     =   8655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   12  'No Drop
   Picture         =   "Principal.frx":0000
   ScaleHeight     =   5895
   ScaleWidth      =   8655
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000006&
      Caption         =   "Financiamento"
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   4
      ToolTipText     =   "Lançamento de Financiamento"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000006&
      Caption         =   "Movimentação de Caixa"
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   8
      ToolTipText     =   "Movimentação de Caixa"
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000006&
      Caption         =   "Moviment. Financeiras"
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   7
      ToolTipText     =   "Movimentação de Financeiras"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000006&
      Caption         =   "Movimentação Clientes"
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   6
      ToolTipText     =   "Movimentação de Clientes"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000006&
      Caption         =   "Prop. Devolução Doc."
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   5
      ToolTipText     =   "Propriedades de Devolução de Cheques e Pagamentos de taxas"
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000006&
      Caption         =   "Troca de Cheques"
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   3
      ToolTipText     =   "Lançamento de Troca de Cheques"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000006&
      Caption         =   "Lançamentos"
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   2
      ToolTipText     =   "Lançamentos de Créditos e Débitos"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000006&
      Caption         =   "Cadastro de Clientes"
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   10
      ToolTipText     =   "Cadastro de Clientes"
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000006&
      Caption         =   "Cadastro de Financeiras"
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   9
      ToolTipText     =   "Cadastro de Financeiras"
      Top             =   4440
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000006&
      Caption         =   "Calculo Financiamento"
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   1
      ToolTipText     =   "Simples calculo para financiamento"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000006&
      Caption         =   "Sair"
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
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      ToolTipText     =   "Abandona o sistema Brusk Factoring"
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000006&
      Caption         =   "Calculo Troca Cheque"
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
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   0
      ToolTipText     =   "Simples calculo para troca de cheque"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   2  'Horizontal Line
      Height          =   255
      Left            =   240
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   2  'Horizontal Line
      Height          =   255
      Left            =   240
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   2  'Horizontal Line
      Height          =   255
      Left            =   240
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillStyle       =   2  'Horizontal Line
      Height          =   255
      Left            =   240
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   240
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command10_Click()
FINANCEIRASR.Show
End Sub

Private Sub Command11_Click()
MOVCAIXA.Show
End Sub

Private Sub Command12_Click()
Calculosl.Show
End Sub

Private Sub Command2_Click()
Calculotroca.Show
End Sub

Private Sub Command3_Click()
Calculos.Show
End Sub

Private Sub Command4_Click()
Financeiras.Show
End Sub

Private Sub Command5_Click()
CLIENTES.Show
End Sub

Private Sub Command6_Click()
Lançamentos.Show
End Sub

Private Sub Command7_Click()
Calculotrocal.Show
End Sub

Private Sub Command8_Click()
DEVOLUÇÃO.Show
End Sub

Private Sub Command9_Click()
CLIENTESR.Show
End Sub
