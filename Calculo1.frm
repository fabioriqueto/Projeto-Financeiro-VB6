VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Calculos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Calculo para Financiamento                                  ***Megatron***"
   ClientHeight    =   5715
   ClientLeft      =   1470
   ClientTop       =   1755
   ClientWidth     =   6360
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6360
   Begin VB.Frame Frame1 
      Caption         =   "Resultado do Financiamento"
      Height          =   2055
      Left            =   120
      TabIndex        =   27
      Top             =   3600
      Visible         =   0   'False
      Width           =   6135
      Begin MSMask.MaskEdBox MaskEdBox17 
         Height          =   375
         Left            =   3960
         TabIndex        =   47
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
      Begin MSMask.MaskEdBox MaskEdBox16 
         Height          =   255
         Left            =   1320
         TabIndex        =   45
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
      Begin MSMask.MaskEdBox MaskEdBox15 
         Height          =   255
         Left            =   4320
         TabIndex        =   43
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
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   255
         Left            =   4320
         TabIndex        =   42
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
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   255
         Left            =   4320
         TabIndex        =   41
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
      Begin MSMask.MaskEdBox MaskEdBox12 
         Height          =   255
         Left            =   4320
         TabIndex        =   40
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
      Begin MSMask.MaskEdBox MaskEdBox11 
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   1680
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   1320
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   255
         Left            =   720
         TabIndex        =   33
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   255
         Left            =   720
         TabIndex        =   32
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   327680
         Enabled         =   0   'False
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
         Caption         =   "Valor Total:"
         Height          =   255
         Left            =   3000
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "CPMF de 0.30:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Valor"
         Height          =   255
         Left            =   3840
         TabIndex        =   39
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Valor"
         Height          =   255
         Left            =   3840
         TabIndex        =   38
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Valor"
         Height          =   255
         Left            =   3840
         TabIndex        =   37
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Valor"
         Height          =   255
         Left            =   3840
         TabIndex        =   36
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Venc. 4"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Venc. 3"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Venc. 2"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Venc. 1"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   10
      Top             =   3600
      Width           =   6135
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datas Vencimentos"
      Height          =   1815
      Left            =   1800
      TabIndex        =   17
      Top             =   1680
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "AUTO"
         Height          =   1335
         Left            =   4080
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   327680
         AllowPrompt     =   -1  'True
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   255
         Left            =   960
         TabIndex        =   5
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   960
         TabIndex        =   6
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
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   255
         Left            =   960
         TabIndex        =   7
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
      Begin VB.Label Label7 
         Caption         =   "4 parcela"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "3 parcela"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "2 parcela"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "1 parcela"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de taxa"
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton Option11 
         Caption         =   "Taxa Postecipada"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Taxa Antecipada"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Taxa Mensal"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Forma Pagamento"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
      Begin VB.OptionButton Option8 
         Caption         =   "4 parcelas"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "3 parcelas"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option6 
         Caption         =   "2 parcelas"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "1 parcela"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Dados Para Calculo"
      Height          =   1455
      Left            =   2160
      TabIndex        =   22
      Top             =   120
      Width           =   4095
      Begin MSMask.MaskEdBox MaskEdBox7 
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
      Begin MSMask.MaskEdBox MaskEdBox6 
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   327680
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         Format          =   "dddddd"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Juros:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Val. Finan.:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   855
      End
   End
End
Attribute VB_Name = "Calculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim juros1 As Integer
Dim dias As Integer

Private Sub Command1_Click()
If Option5.Value = True Then
   dias = (CDate(MaskEdBox2) + 1) - CDate(MaskEdBox1)
End If
If Option6.Value = True Then
   dias = (CDate(MaskEdBox3) + 1) - CDate(MaskEdBox1)
End If
If Option7.Value = True Then
   dias = (CDate(MaskEdBox4) + 1) - CDate(MaskEdBox1)
End If
If Option8.Value = True Then
   dias = (CDate(MaskEdBox5) + 1) - CDate(MaskEdBox1)
End If
MaskEdBox8 = MaskEdBox2
MaskEdBox9 = MaskEdBox3
MaskEdBox10 = MaskEdBox4
MaskEdBox11 = MaskEdBox5
Command1.Visible = False
If Option5.Value = True Then
   If Option10.Value = True Then
      juros1 = (MaskEdBox6 / 30) * dias
      MaskEdBox17 = (MaskEdBox7 / ((100 - juros1) / 100)) / 0.997
      MaskEdBox16 = MaskEdBox17 * 0.003
   ElseIf Option11.Value = True Then
      juros1 = (MaskEdBox6 / 30) * dias
      MaskEdBox17 = (MaskEdBox7 * ((juros1 / 100) + 1)) / 0.997
      MaskEdBox16 = MaskEdBox17 * 0.003
   End If
   MaskEdBox12 = MaskEdBox17
   Label8.Visible = True
   Label12.Visible = True
   MaskEdBox8.Visible = True
   MaskEdBox12.Visible = True
   Label9.Visible = False
   Label10.Visible = False
   Label11.Visible = False
   Label13.Visible = False
   Label14.Visible = False
   Label15.Visible = False
   MaskEdBox9.Visible = False
   MaskEdBox10.Visible = False
   MaskEdBox11.Visible = False
   MaskEdBox13.Visible = False
   MaskEdBox14.Visible = False
   MaskEdBox15.Visible = False
Else
   If Option6.Value = True Then
      dias = (CDate(MaskEdBox3) + 1) - CDate(MaskEdBox1)
      juros1 = (MaskEdBox6 / 30) * dias
      MaskEdBox17 = (MaskEdBox7 * ((juros1 / 100) + 1)) / 0.997
      MaskEdBox16 = MaskEdBox17 * 0.003
      MaskEdBox12 = MaskEdBox17 / 2
      MaskEdBox13 = MaskEdBox17 / 2
      Label8.Visible = True
      Label12.Visible = True
      MaskEdBox8.Visible = True
      MaskEdBox12.Visible = True
      Label9.Visible = True
      Label10.Visible = False
      Label11.Visible = False
      Label13.Visible = True
      Label14.Visible = False
      Label15.Visible = False
      MaskEdBox9.Visible = True
      MaskEdBox10.Visible = False
      MaskEdBox11.Visible = False
      MaskEdBox13.Visible = True
      MaskEdBox14.Visible = False
      MaskEdBox15.Visible = False
   End If
   If Option7.Value = True Then
      dias = (CDate(MaskEdBox4) + 1) - CDate(MaskEdBox1)
      juros1 = (MaskEdBox6 / 30) * dias
      MaskEdBox17 = (MaskEdBox7 * ((juros1 / 100) + 1)) / 0.997
      MaskEdBox16 = MaskEdBox17 * 0.003
      MaskEdBox12 = MaskEdBox17 / 3
      MaskEdBox13 = MaskEdBox17 / 3
      MaskEdBox14 = MaskEdBox17 / 3
      Label8.Visible = True
      Label12.Visible = True
      MaskEdBox8.Visible = True
      MaskEdBox12.Visible = True
      Label9.Visible = True
      Label10.Visible = True
      Label11.Visible = False
      Label13.Visible = True
      Label14.Visible = True
      Label15.Visible = False
      MaskEdBox9.Visible = True
      MaskEdBox10.Visible = True
      MaskEdBox11.Visible = False
      MaskEdBox13.Visible = True
      MaskEdBox14.Visible = True
      MaskEdBox15.Visible = False
   End If
   If Option8.Value = True Then
      dias = (CDate(MaskEdBox5) + 1) - CDate(MaskEdBox1)
      juros1 = (MaskEdBox6 / 30) * dias
      MaskEdBox17 = (MaskEdBox7 * ((juros1 / 100) + 1)) / 0.997
      MaskEdBox16 = MaskEdBox17 * 0.003
      MaskEdBox12 = MaskEdBox17 / 4
      MaskEdBox13 = MaskEdBox17 / 4
      MaskEdBox14 = MaskEdBox17 / 4
      MaskEdBox15 = MaskEdBox17 / 4
      Label8.Visible = True
      Label12.Visible = True
      MaskEdBox8.Visible = True
      MaskEdBox12.Visible = True
      Label9.Visible = True
      Label10.Visible = True
      Label11.Visible = True
      Label13.Visible = True
      Label14.Visible = True
      Label15.Visible = True
      MaskEdBox9.Visible = True
      MaskEdBox10.Visible = True
      MaskEdBox11.Visible = True
      MaskEdBox13.Visible = True
      MaskEdBox14.Visible = True
      MaskEdBox15.Visible = True
   End If
End If


Frame1.Visible = True

End Sub

Private Sub Command2_Click()
wdia1 = Day(CDate(MaskEdBox2))
wmes1 = Month(CDate(MaskEdBox2))
wano1 = Year(CDate(MaskEdBox2))
wmes1 = Val(wmes1) + 1
If Val(wmes1) > 12 Then
   wmes1 = 1
   wano1 = Val(wano1) + 1
End If
MaskEdBox3 = Str(wdia1) + "-" + Str(wmes1) + "-" + Str(wano1)
wteste = MaskEdBox3.FormattedText
If Left(Trim(wteste), 5) = "31- 2" Then MaskEdBox3 = "28-2-" + Str(wano1)
If Left(Trim(wteste), 5) = "30- 2" Then MaskEdBox3 = "28-2-" + Str(wano1)
If Left(Trim(wteste), 5) = "29- 2" Then MaskEdBox3 = "28-2-" + Str(wano1)

wdia2 = Day(CDate(MaskEdBox2))
wmes2 = Month(CDate(MaskEdBox3))
wano2 = Year(CDate(MaskEdBox3))
wmes2 = Val(wmes2) + 1
If Val(wmes2) > 12 Then
   wmes2 = 1
   wano2 = Val(wano2) + 1
End If
MaskEdBox4 = Str(wdia2) + "-" + Str(wmes2) + "-" + Str(wano2)
wteste = MaskEdBox4.FormattedText
If Left(Trim(wteste), 5) = "31- 2" Then MaskEdBox4 = "28-2-" + Str(wano2)
If Left(Trim(wteste), 5) = "30- 2" Then MaskEdBox4 = "28-2-" + Str(wano2)
If Left(Trim(wteste), 5) = "29- 2" Then MaskEdBox4 = "28-2-" + Str(wano2)

wdia3 = Day(CDate(MaskEdBox2))
wmes3 = Month(CDate(MaskEdBox4))
wano3 = Year(CDate(MaskEdBox4))
wmes3 = Val(wmes3) + 1
If Val(wmes3) > 12 Then
   wmes3 = 1
   wano3 = Val(wano3) + 1
End If
MaskEdBox5 = Str(wdia3) + "-" + Str(wmes3) + "-" + Str(wano3)
wteste = MaskEdBox5.FormattedText
If Left(Trim(wteste), 5) = "31- 2" Then MaskEdBox5 = "1-3-" + Str(wano3)
If Left(Trim(wteste), 5) = "30- 2" Then MaskEdBox5 = "1-3-" + Str(wano3)
If Left(Trim(wteste), 5) = "29- 2" Then MaskEdBox5 = "1-3-" + Str(wano3)



End Sub

Private Sub Form_Activate()
MaskEdBox1 = Date
MaskEdBox2 = Date
MaskEdBox3 = Date
MaskEdBox4 = Date
MaskEdBox5 = Date
MaskEdBox6 = 0
MaskEdBox7 = 0
End Sub

Private Sub MaskEdBox1_Change()
If MaskEdBox1.FormattedText = "" Then MaskEdBox1 = Date
Command1.Visible = True
Frame1.Visible = False

End Sub

Private Sub MaskEdBox2_Change()
If MaskEdBox2.FormattedText = "" Then MaskEdBox2 = Date
Command1.Visible = True
Frame1.Visible = False

End Sub

Private Sub MaskEdBox3_Change()
If MaskEdBox3.FormattedText = "" Then MaskEdBox3 = Date
Command1.Visible = True
Frame1.Visible = False

End Sub

Private Sub MaskEdBox4_Change()
If MaskEdBox4.FormattedText = "" Then MaskEdBox4 = Date
Command1.Visible = True
Frame1.Visible = False

End Sub

Private Sub MaskEdBox5_Change()
If MaskEdBox5.FormattedText = "" Then MaskEdBox5 = Date
Command1.Visible = True
Frame1.Visible = False

End Sub

Private Sub MaskEdBox6_Change()
Command1.Visible = True
Frame1.Visible = False
If MaskEdBox6.FormattedText = "" Then MaskEdBox6 = 0

End Sub

Private Sub MaskEdBox7_Change()
Command1.Visible = True
Frame1.Visible = False
If MaskEdBox7.FormattedText = "" Then MaskEdBox7 = 0
End Sub

Private Sub Option10_Click()
Command1.Visible = True
Frame1.Visible = False

End Sub

Private Sub Option11_Click()
Command1.Visible = True
Frame1.Visible = False

End Sub

Private Sub Option5_Click()
Option9.Enabled = False
Option10.Enabled = True
Option11.Enabled = True
Option10.Value = True
MaskEdBox2.Enabled = True
MaskEdBox3.Enabled = False
MaskEdBox4.Enabled = False
MaskEdBox5.Enabled = False
Command1.Visible = True
Frame1.Visible = False
End Sub

Private Sub Option6_Click()
Option9.Enabled = True
Option10.Enabled = False
Option11.Enabled = False
Option9.Value = True
MaskEdBox2.Enabled = True
MaskEdBox3.Enabled = True
MaskEdBox4.Enabled = False
MaskEdBox5.Enabled = False
Command1.Visible = True
Frame1.Visible = False
End Sub

Private Sub Option7_Click()
Option9.Enabled = True
Option10.Enabled = False
Option11.Enabled = False
Option9.Value = True
MaskEdBox2.Enabled = True
MaskEdBox3.Enabled = True
MaskEdBox4.Enabled = True
MaskEdBox5.Enabled = False
Command1.Visible = True
Frame1.Visible = False
End Sub

Private Sub Option8_Click()
Option9.Enabled = True
Option10.Enabled = False
Option11.Enabled = False
Option9.Value = True
MaskEdBox2.Enabled = True
MaskEdBox3.Enabled = True
MaskEdBox4.Enabled = True
MaskEdBox5.Enabled = True
Command1.Visible = True
Frame1.Visible = False
End Sub

Private Sub Option9_Click()
Command1.Visible = True
Frame1.Visible = False

End Sub
