VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmAbrir 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Selecione a Base da Dados Access"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "frmAbrir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picSair 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3330
      MouseIcon       =   "frmAbrir.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "frmAbrir.frx":114C
      ScaleHeight     =   480
      ScaleWidth      =   1455
      TabIndex        =   15
      Top             =   4485
      Width           =   1455
      Begin VB.Label lblSair 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   570
         MouseIcon       =   "frmAbrir.frx":1A4B
         TabIndex        =   16
         Top             =   150
         Width           =   765
      End
   End
   Begin VB.PictureBox picOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   1
      Left            =   1425
      MouseIcon       =   "frmAbrir.frx":1D55
      MousePointer    =   99  'Custom
      Picture         =   "frmAbrir.frx":205F
      ScaleHeight     =   510
      ScaleWidth      =   990
      TabIndex        =   13
      Top             =   4485
      Width           =   990
      Begin VB.Label lblOk 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   585
         MouseIcon       =   "frmAbrir.frx":295E
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   165
         Width           =   255
      End
   End
   Begin VB.ListBox lstSelecao 
      Height          =   3375
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   2535
   End
   Begin VB.ListBox lstCampos 
      Height          =   3375
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtBdd 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
   Begin VB.ComboBox cmbTabelas 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog Controle 
      Left            =   5055
      Top             =   4425
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblCaminho 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   195
      Left            =   135
      TabIndex        =   18
      Top             =   4350
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   210
      Left            =   2820
      TabIndex        =   17
      Top             =   3750
      Width           =   180
   End
   Begin VB.Label lblAbrir 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   2775
      MouseIcon       =   "frmAbrir.frx":2C68
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "Remover campo(s) selecionado(s)"
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label lblAbrir 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   2775
      MouseIcon       =   "frmAbrir.frx":2F72
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "Remover todos os campos"
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label lblAbrir 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   2775
      MouseIcon       =   "frmAbrir.frx":327C
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Adcionar campo(s) selecionado(s)"
      Top             =   1365
      Width           =   240
   End
   Begin VB.Label lblAbrir 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   2775
      MouseIcon       =   "frmAbrir.frx":3586
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Adcionar todos os campos"
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Tabelas:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Campos do Formulário:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   2265
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Campos da Tabela:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1845
   End
   Begin VB.Label lblAbrir 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   2790
      MouseIcon       =   "frmAbrir.frx":3890
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Abrir banco de dados access"
      Top             =   360
      Width           =   240
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Banco de Dados:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1605
   End
End
Attribute VB_Name = "frmAbrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BDD As DAO.Database
Dim i As Integer
Dim xParte As Byte, xCampos As Byte
Dim buf As String * 256
Dim length As Long
Dim Caminho As String, Bloco As String, NomeTab As String, Copia As String
Dim RetVal
Dim xCont As Byte, COntador As Byte
Dim ContaCampo As Integer

Private Sub cmbTabelas_Click()
Call AddCampos(cmbTabelas.Text)
End Sub
Private Sub Form_Load()
If Idioma = "BR" Then
Me.Caption = "Selecione a Base de Dados Access"
Label8.Caption = "Base de Dados"
Label3.Caption = "Tabelas"
Label1.Caption = "Campos da Tabela:"
Label2.Caption = "Campos do Formulário:"
lblSair.Caption = "Cancelar"
Else
Me.Caption = "Select DataBase Access "
Label8.Caption = "Data Base"
Label3.Caption = "Tables"
Label1.Caption = "Field of the Table:"
Label2.Caption = "Field of the Form:"
lblSair.Caption = "Cancel"
End If

abribanco
End Sub

Private Sub lblAbrir_Click(Index As Integer)
On Error Resume Next
 T = cmbTabelas.Text

Select Case Index
Case 0
 lstCampos.Clear
 lstSelecao.Clear
 For Each vgCp In BDD.TableDefs(T).Fields           'para cada campo da tabela
  If Left$(vgCp.ValidationText, 1) <> "I" And _
   InStr(vgCp.Name, "~") = 0 Then                    'se não for invisível ou de controle
    X$ = vgCp.Name
   'X$ = RPad$(vgCp.Name, 25, " ") + " " + _
        Mid$(vgTipoCp$, vgCp.Type * 8 - 7, 8) + _
        LPad$(Str$(vgCp.Size), 4, " ")               'monta nome, tipo, tamanho
   lstSelecao.AddItem X$                             'e adiciona na lista
  End If
 Next
Case 1
    If lstCampos.Text = "" Then: Exit Sub
    lstSelecao.AddItem lstCampos.Text
    i = lstCampos.Text
    lstCampos.RemoveItem lstCampos.ListIndex
    lstCampos.ListIndex = lstCampos.ListIndex + 1
    'Label19 = Label19 + 1
Case 2
    abribanco
Case 3
AddCampos (cmbTabelas.Text)
Case 4
    If lstSelecao.Text = "" Then: Exit Sub
    lstCampos.AddItem lstSelecao.Text
    lstSelecao.RemoveItem lstSelecao.ListIndex
    lstSelecao.ListIndex = lstSelecao.ListIndex + 1
    'Label19 = Label19 - 1
End Select

End Sub

Private Sub lblOk_Click(Index As Integer)
ok
End Sub

Private Sub lblSair_Click()
Unload Me
End Sub

Private Sub picOk_Click(Index As Integer)
ok
End Sub

Private Sub picSair_Click()
Unload Me
End Sub

Private Sub EncheNomesTabs()
On Error GoTo Abrir
If lblCaminho.Caption = "" Then: Exit Sub
 Dim T As TableDef, Q As QueryDef                       'dimensiona locais
 Set BDD = Workspaces(0).OpenDatabase(lblCaminho)
 cmbTabelas.Clear                                       'enche lista com os nomes das tabelas
 For Each T In BDD.TableDefs                            'para cada tabela da coleção,
  If (T.Attributes And dbSystemObject) = 0 And _
     InStr(T.Name, "~") = 0 Then                        'se não for de sistema ou de consulta,
   cmbTabelas.AddItem T.Name                            'coloca na lista
  End If
 Next

 cmbTabelas.ListIndex = 0                               'seleciona a 1a...
 'lstConsultas.Clear                                       'enche lista com os nomes das queries de consulta
 'For Each Q In BDD.QueryDefs                            'para cada query na coleção,
 ' If InStr(Q.Name, "~") = 0 Then                        'se for de consulta,
 '  lstConsultas.AddItem Q.Name                          'adiciona à lista
 ' End If
 'Next
 'If lstConsultas.ListCount > 0 Then                     'se tiver pelo menos uma,
 ' lstConsultas.ListIndex = 0                            'seleciona a 1a...
 'End If
Exit Sub
Abrir:
If Err = 3024 Then
MsgBox "Base de dados não reconhecida!!" & Chr(10) & _
"Selecione uma Base de Dados Válida!!", vbCritical
Exit Sub
End If
End Sub


Public Sub abribanco()
On Error Resume Next
Controle.Filter = "Arquivos MSAccess (*.mdb)|*.mdb|"
Controle.FilterIndex = 1
Controle.ShowOpen
lblCaminho.Caption = Controle.Filename
EncheNomesTabs
End Sub

Public Sub AddCampos(T As String)
 Label19 = 0
 lstCampos.Clear
 lstSelecao.Clear
 For Each vgCp In BDD.TableDefs(T).Fields           'para cada campo da tabela
  If Left$(vgCp.ValidationText, 1) <> "I" And _
   InStr(vgCp.Name, "~") = 0 Then                    'se não for invisível ou de controle
    X$ = vgCp.Name
   'X$ = RPad$(vgCp.Name, 25, " ") + " " + _
        Mid$(vgTipoCp$, vgCp.Type * 8 - 7, 8) + _
        LPad$(Str$(vgCp.Size), 4, " ")               'monta nome, tipo, tamanho
   lstCampos.AddItem X$                             'e adiciona na lista
   Label19 = Label19 + 1
  End If
 Next

End Sub

'enche as listas com as tabelas
Private Sub EncheListas(T As String)
 Dim vgInd As Index, vgCp As Field, i As Integer, _
     X As String, vgTipoCp As String                 'dimensiona locais
 vgTipoCp$ = LoadResString(3110)                     'string com os tipos de campos
 lstCampos.Clear                                     'limpa a lista de campos e
 'lstIndices.Clear                                    'de índices

 'coloca numero de registro da tabela
 labNReg.Caption = Mid$(Str$(Banco.TableDefs(T).RecordCount), 2) + LoadResString(3050)

 ' enche lstCampos com os nomes dos campos das tabelas
 For Each vgCp In Banco.TableDefs(T).Fields           'para cada campo da tabela
  If Left$(vgCp.ValidationText, 1) <> "I" And _
   InStr(vgCp.Name, "~") = 0 Then                    'se não for invisível ou de controle
   X$ = vgCp.Name
   X$ = RPad$(vgCp.Name, 25, " ") + " " + _
        Mid$(vgTipoCp$, vgCp.Type * 8 - 7, 8) + _
        LPad$(Str$(vgCp.Size), 4, " ")               'monta nome, tipo, tamanho
   lstCampos.AddItem X$                              'e adiciona na lista
  End If
 Next

 ' enche lstIndices com nomes dos indices das tabelas
 For Each vgInd In Banco.TableDefs(T).Indexes         'para cada índice da tabela
  If Not vgInd.Foreign And InStr(vgInd.Name, "~") = 0 Then 'se não for estranjeiro nem de consulta
   X$ = vgInd.Name
   'X$ = RPad$(vgInd.Name, 18, " ") + " " + _
        RPad$(vgInd.Fields, 34, " ") + " " + _
        IIf(vgInd.Primary, LoadResString(3120), LoadResString(3130))  'monta nome, campos
   If vgInd.Unique Then X$ = X$ + LoadResString(3140)                'e tipo e
   lstIndices.AddItem X$                                             'adiciona na lista
  End If
 Next
End Sub

Public Sub ok()
On Error Resume Next
If Idioma = "BR" Then
If lstSelecao.ListCount = 0 Then: MsgBox "Selecione os campos!", vbInformation: Exit Sub
Else
If lstSelecao.ListCount = 0 Then: MsgBox "Select the fields!", vbInformation: Exit Sub
End If
    
    frmMaster.lstCampos.Clear
    For COntador = 0 To Label19
    If lstSelecao.Text = "" Then: GoTo prox
    OpenDB "detalhes"
    
'//---------Inclui itens na tabela
    TbDet.AddNew
    TbDet("d_nome") = lstSelecao.Text
    TbDet("d_rotulo") = lstSelecao.Text
    TbDet("d_valorInicial") = ""
    TbDet("d_tamaho") = 50
    TbDet("d_maximo") = 60
    TbDet("d_senha") = "N"
    TbDet("d_tipo") = "Campo de Texto"
    TbDet("d_nulo") = "S"
    TbDet("d_exibir") = "S"
    TbDet.Update

    CloseDB
'//-----------Fim da inclusão

    frmMaster.lstCampos.AddItem lstSelecao.Text
    Me.lstSelecao.RemoveItem lstSelecao.ListIndex
prox:
    Me.lstSelecao.ListIndex = lstSelecao.ListIndex + 1

    Next
    frmMaster.lblNomeTabela.Caption = cmbTabelas.Text
    frmMaster.txtBdd.Text = lblCaminho.Caption
Unload Me
End Sub


Public Sub CarregaStrings()

End Sub
