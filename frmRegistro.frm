VERSION 5.00
Begin VB.Form frmRegistro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Obtenção da chave de registro..."
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "C a n c e l a"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E n v i a r"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3360
      Width           =   3750
   End
   Begin VB.TextBox txtSerie 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2904
      Width           =   3750
   End
   Begin VB.TextBox txtEstado 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   2448
      Width           =   3750
   End
   Begin VB.TextBox txtCidade 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   1992
      Width           =   3750
   End
   Begin VB.TextBox txtemail 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   1536
      Width           =   3750
   End
   Begin VB.TextBox txtNome 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   3750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Código do programa:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   6
      Top             =   3435
      Width           =   1995
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "N° de série.......:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   2979
      Width           =   1995
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Estado............:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   2523
      Width           =   1995
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cidade............:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   2067
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail............:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   1611
      Width           =   1995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nome..............:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   1155
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   $"frmRegistro.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   6045
   End
End
Attribute VB_Name = "frmRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sucesso As Integer

If Trim(txtemail.Text) = "" Then
   MsgBox "Informe o seu endereço de E-mail", vbCritical, "Dados Incompletos..."
Else
Dim sBodyText As String
   sBodyText = sBodyText & "Nome: " & UCase(txtNome.Text) & "     " & vbCrLf
   sBodyText = sBodyText & "E-Mail: " & LCase(txtemail.Text) & "     " & vbCrLf
   sBodyText = sBodyText & "Cidade: " & UCase(txtCidade.Text) & "     " & vbCrLf
   sBodyText = sBodyText & "Estado: " & UCase(txtEstado.Text) & "     " & vbCrLf
   sBodyText = sBodyText & "N° de Série: " & txtSerie.Text & "     " & vbCrLf
   sBodyText = sBodyText & "Código: " & txtCodigo.Text & "     " & vbCrLf

   site = "mailto:" & Trim("pagemaster@com4.com.br") & "?Subject=Chave para registro do PageMaster&Body=" & sBodyText

   successo = ShellToBrowser(Me, site, 0)
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtSerie.Text = frmMaster.txtSerie.Text
txtCodigo.Text = frmMaster.txtCodRegistro.Text
End Sub

Function ShellToBrowser(Frm As Form, ByVal URL, ByVal WindowStyle)
    
    Dim api As Integer
    api = ShellExecute(Frm.hwnd, "open", URL, "", App.Path, WindowStyle)
 
    'verifica o valor retornado
    If api < 31 Then
        'codigo de erro da api
        MsgBox App.Title & " O seu navegador esta com problemas. " & _
          "Verifique se o seu navegador esta corretamente instalado." & _
          "(Error" & Format(api) & ")", 48, "Navegador Indisponivel"
        ShellToBrowser = False
    ElseIf api = 32 Then
        'arquivo sem associação
        MsgBox App.Title & " não foi possível encontrar uma associação para o arquivo " & _
          URL & " no seu seistema. Verifique o seu Navegador padrão... ", 48, "Navegador indisponivel"
        ShellToBrowser = False
    Else
        'funcionou
        ShellToBrowser = True
    End If
    
End Function


