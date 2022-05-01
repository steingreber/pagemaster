VERSION 5.00
Begin VB.Form frmEditar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editar campo do formulário..."
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7500
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picSair 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4350
      MouseIcon       =   "frmEditar.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmEditar.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   1365
      TabIndex        =   4
      Top             =   1485
      Width           =   1365
      Begin VB.Label lblSair 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar"
         Height          =   195
         Left            =   525
         MouseIcon       =   "frmEditar.frx":0C09
         TabIndex        =   7
         Top             =   135
         Width           =   765
      End
   End
   Begin VB.PictureBox picCriaPaginaOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2265
      MouseIcon       =   "frmEditar.frx":0F13
      MousePointer    =   99  'Custom
      Picture         =   "frmEditar.frx":121D
      ScaleHeight     =   510
      ScaleWidth      =   945
      TabIndex        =   3
      Top             =   1470
      Width           =   945
      Begin VB.Label lblCriaPaginaOk 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         Height          =   195
         Left            =   600
         MouseIcon       =   "frmEditar.frx":1B1C
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   180
         Width           =   255
      End
   End
   Begin VB.CheckBox ckObrigatorio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Este campo é com preenchimento obrigatório!!"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   195
      TabIndex        =   1
      Top             =   1035
      Width           =   4575
   End
   Begin VB.CheckBox ckSenha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Campo de Senha..."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5325
      TabIndex        =   2
      Top             =   1035
      Width           =   2040
   End
   Begin VB.TextBox txtCampo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   195
      TabIndex        =   0
      Top             =   450
      Width           =   7185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   105
      X2              =   7365
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nome do Campo:"
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
      Left            =   195
      TabIndex        =   5
      Top             =   210
      Width           =   1635
   End
End
Attribute VB_Name = "frmEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo errload

    OpenDB "Select * From aspmail Where am_codigo=" & vEdita
    
    txtCampo.Text = TbDet("am_descricao")
    
    If TbDet("am_obrigatorio") = "S" Then
    ckObrigatorio.Value = 1
    Else
    ckObrigatorio.Value = 0
    End If
    
    If TbDet("am_senha") = "S" Then
    ckSenha.Value = 1
    Else
    ckSenha.Value = 0
    End If
    
    vEdita = TbDet("id_aspMail")
    
    CloseDB
    
errload:
If Err = 3021 Then
MsgBox "Selecione um item para editar!", vbExclamation, "Atenção!!"
    Banco.Close
    Set Banco = Nothing
Unload Me
End If

End Sub

Private Sub lblCriaPaginaOk_Click()
gravar
End Sub

Private Sub lblSair_Click()
Unload Me
End Sub

Private Sub picCriaPaginaOk_Click()
gravar
End Sub

Private Sub picSair_Click()
Unload Me
End Sub

Public Sub gravar()
    'Abre BD
    OpenDB "Select * From aspmail Where id_aspMail=" & vEdita

    TbDet.Edit
    TbDet("am_descricao") = txtCampo.Text
    
    If ckObrigatorio.Value = 1 Then
    TbDet("am_obrigatorio") = "S"
    Else
    TbDet("am_obrigatorio") = "N"
    End If
    
    If ckSenha.Value = 1 Then
    TbDet("am_senha") = "S"
    Else
    TbDet("am_senha") = "N"
    End If
    TbDet.Update
    
    CloseDB
    
    frmMaster.lstEtiqsMail.Clear
    OpenDB "aspmail"
    
    Do While Not TbDet.EOF
    frmMaster.lstEtiqsMail.AddItem TbDet("am_descricao")
    TbDet.MoveNext
    Loop
    
    CloseDB
    
    Unload Me
End Sub
