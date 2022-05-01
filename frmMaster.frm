VERSION 5.00
Begin VB.Form frmMaster 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Page Master::."
   ClientHeight    =   6465
   ClientLeft      =   3600
   ClientTop       =   3390
   ClientWidth     =   10395
   Icon            =   "frmMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   6465
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freRegistro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Registre-se..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3060
      Left            =   900
      TabIndex        =   145
      Top             =   2475
      Width           =   5355
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   600
         MouseIcon       =   "frmMaster.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":074C
         ScaleHeight     =   450
         ScaleWidth      =   1275
         TabIndex        =   155
         Top             =   2505
         Width           =   1275
         Begin VB.Label Label41 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Obter a chave"
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
            Left            =   555
            MouseIcon       =   "frmMaster.frx":104B
            MousePointer    =   99  'Custom
            TabIndex        =   156
            Top             =   30
            Width           =   675
         End
      End
      Begin VB.TextBox txtSerie 
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
         Left            =   2475
         Locked          =   -1  'True
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   1110
         Width           =   2475
      End
      Begin VB.PictureBox picRegistro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   3705
         MouseIcon       =   "frmMaster.frx":1355
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":165F
         ScaleHeight     =   450
         ScaleWidth      =   1485
         TabIndex        =   151
         Top             =   2505
         Width           =   1485
         Begin VB.Label lblRegistro 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Registrar"
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
            Left            =   555
            MouseIcon       =   "frmMaster.frx":1F5E
            MousePointer    =   99  'Custom
            TabIndex        =   152
            Top             =   105
            Width           =   780
         End
      End
      Begin VB.TextBox txtChaveRegistro 
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
         Left            =   2505
         TabIndex        =   150
         Top             =   2010
         Width           =   2475
      End
      Begin VB.TextBox txtCodRegistro 
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
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2475
      End
      Begin VB.Label Label40 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Número de Série..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   465
         TabIndex        =   154
         Top             =   1155
         Width           =   1635
      End
      Begin VB.Label Label39 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Chave de liberação..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   465
         TabIndex        =   148
         Top             =   2085
         Width           =   1845
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Código do programa..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   465
         TabIndex        =   147
         Top             =   1605
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Para poder usar o PageMaster entre com as informações do registro..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   165
         TabIndex        =   146
         Top             =   450
         Width           =   4920
      End
   End
   Begin VB.Frame freConfigAspMail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Configurações da Página do AspMail..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   9840
      TabIndex        =   97
      Top             =   3840
      Visible         =   0   'False
      Width           =   9810
      Begin VB.TextBox txtServidor 
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
         TabIndex        =   98
         ToolTipText     =   "Ex.: mail.servidor.com.br"
         Top             =   630
         Width           =   4305
      End
      Begin VB.TextBox txtAssunto 
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
         Left            =   4680
         TabIndex        =   100
         ToolTipText     =   "Ex: E-Mail de Contato do Site"
         Top             =   645
         Width           =   4785
      End
      Begin VB.TextBox txtRementente 
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
         TabIndex        =   102
         ToolTipText     =   "Ex: fulano@servidor.com.br"
         Top             =   1380
         Width           =   4305
      End
      Begin VB.TextBox txtNomeRemet 
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
         Left            =   4680
         TabIndex        =   104
         ToolTipText     =   "Ex: José da Silva"
         Top             =   1395
         Width           =   4785
      End
      Begin VB.TextBox txtDestino 
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
         Left            =   210
         TabIndex        =   106
         ToolTipText     =   "Ex: empresa@servidor.com.br"
         Top             =   2160
         Width           =   4305
      End
      Begin VB.TextBox txtTituloPagina 
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
         TabIndex        =   112
         ToolTipText     =   "Nome que aparecerá na barra de título do navegador"
         Top             =   3810
         Width           =   4305
      End
      Begin VB.TextBox txtNomeArq 
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
         Left            =   4770
         TabIndex        =   114
         ToolTipText     =   "Nome que define a página ASP"
         Top             =   3810
         Width           =   4680
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "GET"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   210
         TabIndex        =   108
         ToolTipText     =   "GET - Suporta no máximo 256 caracteres por campo e é enviada através da barra de endereço"
         Top             =   2895
         Width           =   705
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "POST"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1095
         TabIndex        =   110
         ToolTipText     =   "POST - Não tem limites de caracteres e é enviada de forma oculta."
         Top             =   2895
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Nome do servidor SMTP:"
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
         TabIndex        =   113
         Top             =   390
         Width           =   2370
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Assunto da Mensagem:"
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
         Left            =   4680
         TabIndex        =   111
         Top             =   390
         Width           =   2235
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "E-Mail do remetente da mensagem:"
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
         TabIndex        =   109
         Top             =   1140
         Width           =   3465
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Nome do remetente da mensagem:"
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
         Left            =   4680
         TabIndex        =   107
         Top             =   1140
         Width           =   3420
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "E-Mail para onde a mensagem será enviada:"
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
         TabIndex        =   105
         Top             =   1920
         Width           =   4350
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Título da página"
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
         TabIndex        =   103
         Top             =   3570
         Width           =   1575
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Nome da página ASP:"
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
         Left            =   4770
         TabIndex        =   101
         Top             =   3570
         Width           =   2070
      End
      Begin VB.Label Label35 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Método de envio da Mensagem:"
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
         Left            =   210
         TabIndex        =   99
         Top             =   2685
         Width           =   3060
      End
   End
   Begin VB.Frame freGeradoAsp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Formulário AspMail gerado com sucesso!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5610
      Left            =   9705
      TabIndex        =   115
      Top             =   660
      Visible         =   0   'False
      Width           =   9810
      Begin VB.PictureBox picCriaPaginaOk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   8685
         MouseIcon       =   "frmMaster.frx":2268
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":2572
         ScaleHeight     =   510
         ScaleWidth      =   945
         TabIndex        =   122
         Top             =   5010
         Width           =   945
         Begin VB.Label lblCriaPaginaOk 
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
            Left            =   600
            MouseIcon       =   "frmMaster.frx":2E71
            MousePointer    =   99  'Custom
            TabIndex        =   123
            Top             =   180
            Width           =   255
         End
      End
      Begin VB.PictureBox picEnviarDados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   6412
         MouseIcon       =   "frmMaster.frx":317B
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":3485
         ScaleHeight     =   510
         ScaleWidth      =   1860
         TabIndex        =   120
         Top             =   5010
         Width           =   1860
         Begin VB.Label lblEnviarDados 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Página para enviar dados.."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   555
            MouseIcon       =   "frmMaster.frx":40FC
            MousePointer    =   99  'Custom
            TabIndex        =   121
            Top             =   60
            Width           =   1320
         End
      End
      Begin VB.PictureBox picInserirDados 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   4125
         MouseIcon       =   "frmMaster.frx":4406
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":4710
         ScaleHeight     =   510
         ScaleWidth      =   1890
         TabIndex        =   118
         Top             =   5010
         Width           =   1890
         Begin VB.Label lblInserirDados 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Página para inserir dados.."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   555
            MouseIcon       =   "frmMaster.frx":5387
            MousePointer    =   99  'Custom
            TabIndex        =   119
            Top             =   60
            Width           =   1365
         End
      End
      Begin VB.TextBox txtGeradoP2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   117
         Top             =   210
         Visible         =   0   'False
         Width           =   9660
      End
      Begin VB.TextBox txtGerado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBEDC4&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   116
         Top             =   210
         Width           =   9660
      End
   End
   Begin VB.Frame freForms 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Formulários gerados com sucesso!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5610
      Left            =   8865
      TabIndex        =   132
      Top             =   3030
      Visible         =   0   'False
      Width           =   10215
      Begin VB.PictureBox picTAlteracao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   7440
         MouseIcon       =   "frmMaster.frx":5691
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":599B
         ScaleHeight     =   450
         ScaleWidth      =   1455
         TabIndex        =   143
         Top             =   5025
         Width           =   1455
         Begin VB.Label lblTAlteracao 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Alteração"
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
            Left            =   450
            MouseIcon       =   "frmMaster.frx":6A12
            MousePointer    =   99  'Custom
            TabIndex        =   144
            Top             =   135
            Width           =   810
         End
      End
      Begin VB.PictureBox picTInclusao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   5400
         MouseIcon       =   "frmMaster.frx":6D1C
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":7026
         ScaleHeight     =   450
         ScaleWidth      =   1455
         TabIndex        =   141
         Top             =   5025
         Width           =   1455
         Begin VB.Label lblTInclusao 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Inclusão"
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
            Left            =   465
            MouseIcon       =   "frmMaster.frx":809D
            MousePointer    =   99  'Custom
            TabIndex        =   142
            Top             =   135
            Width           =   720
         End
      End
      Begin VB.PictureBox picTConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   3345
         MouseIcon       =   "frmMaster.frx":83A7
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":86B1
         ScaleHeight     =   450
         ScaleWidth      =   1455
         TabIndex        =   139
         Top             =   5025
         Width           =   1455
         Begin VB.Label lblTConsulta 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Consulta"
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
            Left            =   465
            MouseIcon       =   "frmMaster.frx":9728
            MousePointer    =   99  'Custom
            TabIndex        =   140
            Top             =   135
            Width           =   750
         End
      End
      Begin VB.TextBox txtCriaConexao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBEDC4&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   136
         Top             =   210
         Width           =   10080
      End
      Begin VB.TextBox txtCriaAlteracao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   138
         Top             =   210
         Visible         =   0   'False
         Width           =   10080
      End
      Begin VB.TextBox txtCriaInclusao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   137
         Top             =   210
         Visible         =   0   'False
         Width           =   10080
      End
      Begin VB.TextBox txtCriaConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   135
         Top             =   210
         Visible         =   0   'False
         Width           =   10080
      End
      Begin VB.PictureBox picTConexao 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   1305
         MouseIcon       =   "frmMaster.frx":9A32
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":9D3C
         ScaleHeight     =   450
         ScaleWidth      =   1455
         TabIndex        =   133
         Top             =   5025
         Width           =   1455
         Begin VB.Label lblTConexao 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Conexão"
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
            Left            =   465
            MouseIcon       =   "frmMaster.frx":ADB3
            MousePointer    =   99  'Custom
            TabIndex        =   134
            Top             =   135
            Width           =   765
         End
      End
   End
   Begin VB.Frame freAspMail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Entre com os campos que dejea enviar com o AspMail..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   8520
      TabIndex        =   57
      Top             =   4200
      Visible         =   0   'False
      Width           =   9810
      Begin VB.PictureBox picMail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   3
         Left            =   4890
         MouseIcon       =   "frmMaster.frx":B0BD
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":B3C7
         ScaleHeight     =   510
         ScaleWidth      =   1425
         TabIndex        =   124
         Top             =   4350
         Width           =   1425
         Begin VB.Label lblMail 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Editar Item"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   555
            MouseIcon       =   "frmMaster.frx":BC91
            MousePointer    =   99  'Custom
            TabIndex        =   125
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.Timer Timer1 
         Left            =   7440
         Top             =   4200
      End
      Begin VB.CheckBox ckSenha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Campo de Senha..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5370
         TabIndex        =   60
         Top             =   1080
         Width           =   2040
      End
      Begin VB.CheckBox ckObrigatorio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Este campo é com preenchimento obrigatório!!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   59
         Top             =   1080
         Width           =   4575
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Deseja incluir uma caixa de mensagem em seu fomulário ?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   75
         Top             =   3840
         Width           =   5415
      End
      Begin VB.ListBox lstCamposMail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1920
         ItemData        =   "frmMaster.frx":BF9B
         Left            =   6840
         List            =   "frmMaster.frx":BF9D
         TabIndex        =   73
         Top             =   1680
         Width           =   2610
      End
      Begin VB.ListBox lstVariaveis 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   1920
         ItemData        =   "frmMaster.frx":BF9F
         Left            =   3540
         List            =   "frmMaster.frx":BFA1
         TabIndex        =   71
         Top             =   1680
         Width           =   2610
      End
      Begin VB.ListBox lstEtiqsMail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         ItemData        =   "frmMaster.frx":BFA3
         Left            =   240
         List            =   "frmMaster.frx":BFA5
         TabIndex        =   69
         Top             =   1680
         Width           =   2610
      End
      Begin VB.PictureBox picMail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   3345
         MouseIcon       =   "frmMaster.frx":BFA7
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":C2B1
         ScaleHeight     =   510
         ScaleWidth      =   1425
         TabIndex        =   67
         Top             =   4350
         Width           =   1425
         Begin VB.Label lblMail 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Excluir Tudo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   540
            MouseIcon       =   "frmMaster.frx":CFB0
            MousePointer    =   99  'Custom
            TabIndex        =   68
            Top             =   15
            Width           =   705
         End
      End
      Begin VB.PictureBox picMail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   1
         Left            =   1800
         MouseIcon       =   "frmMaster.frx":D2BA
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":D5C4
         ScaleHeight     =   510
         ScaleWidth      =   1425
         TabIndex        =   65
         Top             =   4350
         Width           =   1425
         Begin VB.Label lblMail 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Excluir Item"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   555
            MouseIcon       =   "frmMaster.frx":DEC3
            TabIndex        =   66
            Top             =   15
            Width           =   660
         End
      End
      Begin VB.PictureBox picMail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   255
         MouseIcon       =   "frmMaster.frx":E1CD
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":E4D7
         ScaleHeight     =   510
         ScaleWidth      =   1425
         TabIndex        =   63
         Top             =   4350
         Width           =   1425
         Begin VB.Label lblMail 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Incluir Item"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   525
            MouseIcon       =   "frmMaster.frx":EDD6
            MousePointer    =   99  'Custom
            TabIndex        =   64
            Top             =   15
            Width           =   615
         End
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
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   7185
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7680
         TabIndex        =   76
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome dos Campos:"
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
         Left            =   6840
         TabIndex        =   74
         Top             =   1440
         Width           =   1845
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Lista das Variaveis:"
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
         Left            =   3540
         TabIndex        =   72
         Top             =   1440
         Width           =   1905
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Rótulos dos Campos:"
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
         Left            =   240
         TabIndex        =   70
         Top             =   1440
         Width           =   2010
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quant. de Campos:"
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
         Left            =   7680
         TabIndex        =   62
         Top             =   360
         Width           =   1815
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
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   1635
      End
      Begin VB.Image Image4 
         Height          =   1875
         Left            =   7590
         Picture         =   "frmMaster.frx":F0E0
         Top             =   105
         Width           =   2205
      End
   End
   Begin VB.Frame freConfCampo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Configuração dos campos do formulário..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   9120
      TabIndex        =   34
      Top             =   2280
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CheckBox ckExibir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Exibir este campo ?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   315
         Left            =   135
         TabIndex        =   131
         ToolTipText     =   "Para não exibir este campo remova o tik..."
         Top             =   1125
         Width           =   1980
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4800
         TabIndex        =   56
         Text            =   "Page Master"
         Top             =   3315
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   4800
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   54
         Text            =   "frmMaster.frx":F959
         Top             =   3315
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox cmbTipoCampo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMaster.frx":F967
         Left            =   165
         List            =   "frmMaster.frx":F971
         TabIndex        =   53
         Text            =   "Campo de Texto"
         ToolTipText     =   "Selecione aqui qual o tipo de campo para este registro..."
         Top             =   4395
         Width           =   1935
      End
      Begin VB.OptionButton optSenha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Não"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   51
         Top             =   3555
         Width           =   855
      End
      Begin VB.OptionButton optSenha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sim"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   50
         Top             =   3555
         Width           =   735
      End
      Begin VB.TextBox txtQuant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3030
         TabIndex        =   45
         Text            =   "60"
         ToolTipText     =   "Digite aqui o valor numerico"
         Top             =   2715
         Width           =   375
      End
      Begin VB.TextBox txtTamCampo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   390
         TabIndex        =   41
         Text            =   "30"
         ToolTipText     =   "Digite aqui o valor numerico"
         Top             =   2715
         Width           =   375
      End
      Begin VB.TextBox txtInicial 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   39
         ToolTipText     =   "Se este campo inicial com algum valor padrão digite aqui..."
         Top             =   1845
         Width           =   7410
      End
      Begin VB.CheckBox ckNulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Este campo aceita valores nulos!!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   315
         Left            =   135
         TabIndex        =   38
         ToolTipText     =   "Para tornar este campo obrigatório remova o tik..."
         Top             =   765
         Width           =   3480
      End
      Begin VB.TextBox txtRot 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   36
         ToolTipText     =   "Digite aqui a legenda do campo..."
         Top             =   480
         Width           =   7395
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Exemplo:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   4800
         TabIndex        =   55
         Top             =   3075
         Width           =   900
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tipo de campo para este registro:"
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
         Left            =   165
         TabIndex        =   52
         Top             =   4155
         Width           =   3315
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Este é um campo de Senhas ?"
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
         Left            =   165
         TabIndex        =   49
         Top             =   3315
         Width           =   2895
      End
      Begin VB.Label lblPaginar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   240
         Index           =   5
         Left            =   3390
         MouseIcon       =   "frmMaster.frx":F995
         MousePointer    =   99  'Custom
         TabIndex        =   48
         ToolTipText     =   "Clique aqui para almentar o valor da quantidade máxima de caracteres no campo... Ou digite na caixa"
         Top             =   2700
         Width           =   210
      End
      Begin VB.Label lblPaginar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   240
         Index           =   4
         Left            =   2820
         MouseIcon       =   "frmMaster.frx":FC9F
         MousePointer    =   99  'Custom
         TabIndex        =   47
         ToolTipText     =   "Clique aqui para diminuir o valor da quantidade máxima de caracteres no campo... Ou digite na caixa"
         Top             =   2700
         Width           =   210
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Quantidade máxima de caracteres:"
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
         Left            =   2805
         TabIndex        =   46
         Top             =   2490
         Width           =   3420
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tamanho deste campo:"
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
         Left            =   165
         TabIndex        =   44
         Top             =   2490
         Width           =   2265
      End
      Begin VB.Label lblPaginar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   240
         Index           =   3
         Left            =   180
         MouseIcon       =   "frmMaster.frx":FFA9
         MousePointer    =   99  'Custom
         TabIndex        =   43
         ToolTipText     =   "Clique aqui para diminuir o valor do tamanho... Ou digite na caixa"
         Top             =   2700
         Width           =   210
      End
      Begin VB.Label lblPaginar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   240
         Index           =   2
         Left            =   750
         MouseIcon       =   "frmMaster.frx":102B3
         MousePointer    =   99  'Custom
         TabIndex        =   42
         ToolTipText     =   "Clique aqui para almentar o valor do tamanho... Ou digite na caixa"
         Top             =   2700
         Width           =   210
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Valor inicial (se necessário)"
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
         Left            =   165
         TabIndex        =   40
         Top             =   1620
         Width           =   2745
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Nome do rótulo para este campo:"
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
         Left            =   135
         TabIndex        =   37
         Top             =   255
         Width           =   3255
      End
   End
   Begin VB.Frame freConfig 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Confirgure a aparência da sua página ASP..."
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
      Height          =   4935
      Left            =   9150
      TabIndex        =   14
      Top             =   990
      Visible         =   0   'False
      Width           =   9810
      Begin VB.TextBox txtBdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   615
         Width           =   9285
      End
      Begin VB.TextBox txtNomeArquivo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5790
         TabIndex        =   127
         Top             =   3525
         Width           =   3690
      End
      Begin VB.TextBox txtAlteracao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5805
         TabIndex        =   31
         Text            =   "Formulário de alteração"
         Top             =   2775
         Width           =   3690
      End
      Begin VB.TextBox txtConsulta 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5790
         TabIndex        =   28
         Text            =   "Formulário de Consulta"
         Top             =   2025
         Width           =   3690
      End
      Begin VB.TextBox txtInclusao 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5790
         TabIndex        =   27
         Text            =   "Formulário de Inclusão"
         Top             =   1320
         Width           =   3690
      End
      Begin VB.ComboBox cmbFonte 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMaster.frx":105BD
         Left            =   255
         List            =   "frmMaster.frx":105CD
         TabIndex        =   24
         Text            =   "Verdana"
         Top             =   2745
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Exibir Bordas da Tabela"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   255
         TabIndex        =   23
         Top             =   2100
         Width           =   4905
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alternar cores da linhas da Tabela"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   255
         TabIndex        =   22
         Top             =   1725
         Width           =   4830
      End
      Begin VB.ComboBox cmbTamanho 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMaster.frx":105FA
         Left            =   255
         List            =   "frmMaster.frx":10607
         TabIndex        =   21
         Text            =   "1"
         Top             =   3435
         Width           =   855
      End
      Begin VB.TextBox txtPaginar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         TabIndex        =   16
         Text            =   "20"
         Top             =   1335
         Width           =   375
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Caminho do Projeto:"
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
         Left            =   210
         TabIndex        =   130
         Top             =   375
         Width           =   1980
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Nome do Arquivo ASP:"
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
         Left            =   5790
         TabIndex        =   128
         Top             =   3300
         Width           =   2175
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Título do Formulário de Alteração:"
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
         Left            =   5805
         TabIndex        =   32
         Top             =   2550
         Width           =   3360
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Título do Formulário de consulta"
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
         Left            =   5805
         TabIndex        =   30
         Top             =   1800
         Width           =   3180
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Título do Formulário de inclusão"
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
         Left            =   5805
         TabIndex        =   29
         Top             =   1095
         Width           =   3165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Fonte"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   26
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tamanho da Fonte"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   25
         Top             =   3210
         Width           =   1815
      End
      Begin VB.Label lblPaginar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   240
         Index           =   1
         Left            =   1500
         MouseIcon       =   "frmMaster.frx":10614
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1320
         Width           =   210
      End
      Begin VB.Label lblPaginar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Height          =   240
         Index           =   0
         Left            =   930
         MouseIcon       =   "frmMaster.frx":1091E
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   1320
         Width           =   210
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Exibir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   255
         TabIndex        =   18
         Top             =   1335
         Width           =   480
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "registros por página.."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1875
         TabIndex        =   17
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Paginação:"
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
         Left            =   255
         TabIndex        =   15
         Top             =   1110
         Width           =   1065
      End
   End
   Begin VB.Frame freCampos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Campos do Fomulário"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   645
      TabIndex        =   33
      Top             =   5535
      Visible         =   0   'False
      Width           =   2340
      Begin VB.ListBox lstCampos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   4320
         ItemData        =   "frmMaster.frx":10C28
         Left            =   105
         List            =   "frmMaster.frx":10C2A
         TabIndex        =   35
         Top             =   570
         Width           =   2130
      End
      Begin VB.Label lblNomeTabela 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "tabela"
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
         Left            =   105
         TabIndex        =   126
         Top             =   255
         Width           =   615
      End
   End
   Begin VB.Frame freOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "O que você deseja criar?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   1320
      TabIndex        =   84
      Top             =   1080
      Width           =   3675
      Begin VB.OptionButton optAsp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Página com formulário ASP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   240
         TabIndex        =   89
         Top             =   495
         Width           =   2730
      End
      Begin VB.OptionButton optMail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enviar informações com AspMail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         TabIndex        =   88
         Top             =   1035
         Width           =   3120
      End
      Begin VB.PictureBox picOk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   2580
         MouseIcon       =   "frmMaster.frx":10C2C
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":10F36
         ScaleHeight     =   510
         ScaleWidth      =   990
         TabIndex        =   86
         Top             =   2040
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
            Index           =   0
            Left            =   585
            MouseIcon       =   "frmMaster.frx":11835
            MousePointer    =   99  'Custom
            TabIndex        =   87
            Top             =   165
            Width           =   255
         End
      End
      Begin VB.OptionButton optCdonts 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enviar informações com CDonts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         TabIndex        =   85
         Top             =   1560
         Width           =   3120
      End
   End
   Begin VB.PictureBox picGeralMail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1815
      ScaleHeight     =   480
      ScaleWidth      =   6735
      TabIndex        =   77
      Top             =   750
      Visible         =   0   'False
      Width           =   6735
      Begin VB.PictureBox picGeraMail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   4695
         MouseIcon       =   "frmMaster.frx":11B3F
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":11E49
         ScaleHeight     =   510
         ScaleWidth      =   1605
         TabIndex        =   82
         Top             =   15
         Width           =   1605
         Begin VB.Label lblGeraMail 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Gerar Form ASP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   525
            MouseIcon       =   "frmMaster.frx":12EC0
            TabIndex        =   83
            Top             =   45
            Width           =   1035
         End
      End
      Begin VB.PictureBox picConfigMail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   2535
         MouseIcon       =   "frmMaster.frx":131CA
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":134D4
         ScaleHeight     =   510
         ScaleWidth      =   1605
         TabIndex        =   80
         Top             =   -15
         Width           =   1605
         Begin VB.Label lblConfigMail 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Configurar Página ASP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   540
            MouseIcon       =   "frmMaster.frx":13DD3
            TabIndex        =   81
            Top             =   30
            Width           =   1110
         End
      End
      Begin VB.PictureBox picDefineMail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   345
         MouseIcon       =   "frmMaster.frx":140DD
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":143E7
         ScaleHeight     =   510
         ScaleWidth      =   1605
         TabIndex        =   78
         Top             =   30
         Width           =   1605
         Begin VB.Label lblDefineMail 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Definição da Página ASP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   525
            MouseIcon       =   "frmMaster.frx":1545E
            TabIndex        =   79
            Top             =   45
            Width           =   1110
         End
      End
   End
   Begin VB.PictureBox picNovo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   210
      MouseIcon       =   "frmMaster.frx":15768
      MousePointer    =   99  'Custom
      Picture         =   "frmMaster.frx":15A72
      ScaleHeight     =   510
      ScaleWidth      =   1305
      TabIndex        =   12
      Top             =   45
      Visible         =   0   'False
      Width           =   1305
      Begin VB.Label lblNovo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Novo Projeto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   555
         MouseIcon       =   "frmMaster.frx":166E9
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.PictureBox picGeral 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1695
      ScaleHeight     =   480
      ScaleWidth      =   6735
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   6735
      Begin VB.PictureBox picAbrir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   0
         MouseIcon       =   "frmMaster.frx":169F3
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":16CFD
         ScaleHeight     =   510
         ScaleWidth      =   1605
         TabIndex        =   10
         Top             =   15
         Width           =   1605
         Begin VB.Label lblAbrir 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Abrir base de dados"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   540
            MouseIcon       =   "frmMaster.frx":175FC
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   60
            Width           =   945
         End
      End
      Begin VB.PictureBox picDefinir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1590
         MouseIcon       =   "frmMaster.frx":17906
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":17C10
         ScaleHeight     =   510
         ScaleWidth      =   1605
         TabIndex        =   8
         Top             =   15
         Width           =   1605
         Begin VB.Label lblDefinir 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Definição da Página ASP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   525
            MouseIcon       =   "frmMaster.frx":18C87
            TabIndex        =   9
            Top             =   45
            Width           =   1110
         End
      End
      Begin VB.PictureBox picConfig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   3435
         MouseIcon       =   "frmMaster.frx":18F91
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":1929B
         ScaleHeight     =   510
         ScaleWidth      =   1605
         TabIndex        =   6
         Top             =   15
         Width           =   1605
         Begin VB.Label lblConfig 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Configurar Página ASP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   525
            MouseIcon       =   "frmMaster.frx":19B9A
            TabIndex        =   7
            Top             =   30
            Width           =   1110
         End
      End
      Begin VB.PictureBox picGerar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   5160
         MouseIcon       =   "frmMaster.frx":19EA4
         MousePointer    =   99  'Custom
         Picture         =   "frmMaster.frx":1A1AE
         ScaleHeight     =   510
         ScaleWidth      =   1605
         TabIndex        =   4
         Top             =   15
         Width           =   1605
         Begin VB.Label lblGerar 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Gerar Form ASP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   525
            MouseIcon       =   "frmMaster.frx":1B225
            TabIndex        =   5
            Top             =   45
            Width           =   1035
         End
      End
   End
   Begin VB.PictureBox picSair 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8535
      MouseIcon       =   "frmMaster.frx":1B52F
      MousePointer    =   99  'Custom
      Picture         =   "frmMaster.frx":1B839
      ScaleHeight     =   480
      ScaleWidth      =   1605
      TabIndex        =   0
      Top             =   60
      Width           =   1605
      Begin VB.Label lblSair 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sair do Page Master !!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   525
         MouseIcon       =   "frmMaster.frx":1C138
         TabIndex        =   1
         Top             =   45
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sobre..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   6330
      TabIndex        =   90
      Top             =   1470
      Width           =   3675
      Begin VB.Image Image1 
         Height          =   480
         Left            =   165
         Picture         =   "frmMaster.frx":1C442
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Nome do Projeto"
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
         Left            =   705
         TabIndex        =   96
         Top             =   570
         Width           =   1635
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sistema para confecção de formulários ASP e envio de e-mail através do AspMail e CDonts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   165
         MouseIcon       =   "frmMaster.frx":1C884
         TabIndex        =   95
         Top             =   855
         Width           =   3435
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Autor:"
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
         Left            =   165
         TabIndex        =   94
         Top             =   1605
         Width           =   600
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alex L. Steingreber"
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
         Left            =   165
         MouseIcon       =   "frmMaster.frx":1CB8E
         TabIndex        =   93
         Top             =   1815
         Width           =   1650
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Página Web:"
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
         Left            =   165
         TabIndex        =   92
         Top             =   2130
         Width           =   1200
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "www.pagemaster.tk"
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
         Left            =   165
         MouseIcon       =   "frmMaster.frx":1CE98
         MousePointer    =   99  'Custom
         TabIndex        =   91
         Top             =   2325
         Width           =   1710
      End
   End
   Begin VB.Image Image3 
      Height          =   2220
      Left            =   -45
      Picture         =   "frmMaster.frx":1D1A2
      Top             =   600
      Width           =   2235
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Informações..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   6105
      Width           =   10305
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   60
      X2              =   10320
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Image Image2 
      Height          =   2415
      Left            =   8295
      Picture         =   "frmMaster.frx":1E005
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Index           =   0
      Begin VB.Menu mnuarq 
         Caption         =   "&Novo Projeto"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuarq 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuarq 
         Caption         =   "&Sair"
         Index           =   2
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "&Configurações"
      Begin VB.Menu mnuConf 
         Caption         =   "Ca&minho"
         Index           =   0
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuConf 
         Caption         =   "&Idioma"
         Index           =   1
         Begin VB.Menu mnuId 
            Caption         =   "Portugues (Brasil)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuId 
            Caption         =   "Inglês (USA)"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
      Index           =   0
      Begin VB.Menu mnucont 
         Caption         =   "Conteúdo"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnucont 
         Caption         =   "Página da &Web"
         Index           =   1
         Shortcut        =   ^W
      End
      Begin VB.Menu mnucont 
         Caption         =   "F A Q..."
         Index           =   2
      End
      Begin VB.Menu mnucont 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnucont 
         Caption         =   "Recomende a um amigo"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vTexto As String, Caminho As String
Dim sql As String, xAbrir As Byte
Dim buf As String * 256
Dim length As Long, vCampo As String

Public Sub AbrirProcura()
frmAbrir.Show vbModal
End Sub

Public Sub prSair()
If Idioma = "BR" Then
If MsgBox(LoadResString(186), vbYesNo + vbQuestion, "Atenção!!") = 6 Then: Unload frmMaster
Else
If MsgBox(LoadResString(265), vbYesNo + vbQuestion, LoadResString(273)) = 6 Then: Unload frmMaster
End If
apagarBanco
End Sub

Private Sub ckExibir_Click()
gravaMudancas
End Sub

Private Sub ckNulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
gravaMudancas
End Sub

Private Sub ckSenha_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case Index
Case 0
If KeyCode = 13 Then
picMail_Click (0)
Timer1.Interval = 100
End If
End Select

End Sub

Private Sub cmbTipoCampo_Click()
If cmbTipoCampo.Text = "Caixa de Texto" Then
optSenha(1).Value = 1
txtTamCampo.Text = 40
txtQuant.Text = 8
If Idioma = "BR" Then
Label11.Caption = "Largura em Caracteres:"
Label10.Caption = "Número de Linhas:"
Else
Label11.Caption = "Width in Characterses:"
Label10.Caption = "Number of Lines:"

End If

Text6.Visible = False
Text5.Visible = True
Else
txtTamCampo.Text = 15
txtQuant.Text = 60
If Idioma = "BR" Then
Label11.Caption = "Tamanho deste campo:"
Label10.Caption = "Quantidade máxima de caracteres:"
Else
Label11.Caption = "Size of this field:"
Label10.Caption = "Maximum amount of characterses:"
End If
Text6.Visible = True
Text5.Visible = False
End If
gravaMudancas
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
 
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Caption = App.Title & " - v" & App.Major & "." & App.Minor & " by " & App.Comments
Label18.Caption = App.Title & " - v" & App.Major & "." & App.Minor
RemoveMenus Me, False, False, False, False, False, True, True

CarregaStrings
abrirFreme

    length = GetPrivateProfileString( _
        "Config", "Idioma", App.Path, _
        buf, Len(buf), App.Path & "\pMaster.ini")
        Idioma = Left$(buf, length)
        
        If Idioma = "BR" Then
            mnuId(0).Checked = True
            mnuId(1).Checked = False
        Else
            mnuId(0).Checked = False
            mnuId(1).Checked = True
        End If

  If ActiveLock.RegisteredUser Then
    freRegistro.Visible = False
  Else
    ActiveLock.Password = DriveSerialNumber("C:") & "hOkAhEy"
    txtSerie.Text = DriveSerialNumber("C:")
    freRegistro.Visible = True
    txtCodRegistro.Text = ActiveLock.SoftwareCode
    txtChaveRegistro.Text = ""
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
apagarBanco
End Sub

Private Sub Label26_Click()
 Dim sucesso As Integer
 Dim site As String
site = "http://www.pagemaster.tk"

sucesso = ShellToBrowser(Me, site, 0)
End Sub

Private Sub Label26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label26.ForeColor = &H80000012 Then
    Label26.ForeColor = vbBlue
  ElseIf Label26.ForeColor = vbBlue Then
    Label26.ForeColor = &H80000012
  End If
End Sub

Private Sub Label41_Click()
obter
End Sub

Private Sub lblAbrir_Click()
AbrirProcura
End Sub

Private Sub lblConfig_Click()
ConfigASP
End Sub

Private Sub lblConfigMail_Click()
pcasp
End Sub

Private Sub lblCriaPaginaOk_Click()
freGeradoAsp.Visible = False
txtGerado.Text = ""
txtGeradoP2.Text = ""
End Sub

Private Sub lblDefineMail_Click()
ldasp
End Sub

Private Sub lblDefinir_Click()
definir
End Sub

Private Sub lblEnviarDados_Click()
txtGerado.Visible = False
txtGeradoP2.Visible = True
End Sub

Private Sub lblGeraMail_Click()
CriarAspMail
End Sub

Private Sub lblGerar_Click()
CriarConexao
End Sub

Private Sub lblInserirDados_Click()
txtGerado.Visible = True
txtGeradoP2.Visible = False
End Sub

Private Sub lblMail_Click(Index As Integer)
On Error GoTo ErroExc
Select Case Index
Case 0
If Idioma = "BR" Then
If txtCampo.Text = "" Then: MsgBox "Informe o nome do campo!", vbInformation: txtCampo.SetFocus: Exit Sub
Else
If txtCampo.Text = "" Then: MsgBox "Inform the name of the field!", vbInformation: txtCampo.SetFocus: Exit Sub
End If

'-------------Procura por nome duplicado------------
    OpenDB "Select am_descricao From aspmail Where am_descricao='" & txtCampo.Text & ":'"
    
    If TbDet("am_descricao") = txtCampo.Text & ":" Then
    If Err = 3021 Then: GoTo erro3021
    
    If Idioma = "BR" Then
    MsgBox "Nome de Campo duplicado!", vbCritical, "Atenção!!"
    Else
    MsgBox "Name of duplicated Field!", vbCritical, "Attention!!"
    End If
    
    CloseDB
    
    Exit Sub
    End If
    
erro3021:
    CloseDB

'----------------------------------------------------
lstEtiqsMail.AddItem txtCampo.Text & ":"
Label28 = Label28 + 1
Dim p As String, i As String, w As String, Y As String, X As String
vCampo = Replace(txtCampo, " ", "", 1)
lstVariaveis.AddItem "v" & lstEtiqsMail.ListCount & Left(UCase(vCampo), 1) & Left(LCase(vCampo), 15)
lstCamposMail.AddItem "t" & lstEtiqsMail.ListCount & Left(UCase(vCampo), 1) & Left(LCase(vCampo), 15)


'------------------------------------ gravar
    OpenDB "aspmail"
    
    TbDet.AddNew
    TbDet("am_codigo") = lstEtiqsMail.ListCount
    TbDet("am_descricao") = txtCampo.Text & ":"
    
    TbDet("am_variavel") = "v" & lstEtiqsMail.ListCount & Left(UCase(vCampo), 1) & Left(LCase(vCampo), 15)
    TbDet("am_campo") = "t" & lstEtiqsMail.ListCount & Left(UCase(vCampo), 1) & Left(LCase(vCampo), 15)

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

'------------------------------fim da gravação

ckObrigatorio.Value = 0
ckSenha.Value = 0
txtCampo.Text = ""
txtCampo.SetFocus

Case 1
If lstVariaveis.ListCount = 0 Then: Exit Sub

lstVariaveis.RemoveItem lstEtiqsMail.ListIndex
lstCamposMail.RemoveItem lstEtiqsMail.ListIndex
lstEtiqsMail.RemoveItem lstEtiqsMail.ListIndex
Label28 = Label28 - 1

Case 2
If lstVariaveis.ListCount = 0 Then: Exit Sub
If Idioma = "BR" Then
If MsgBox("Deseja apagar todos os itens?", vbQuestion + vbYesNo, "Atenção!!") = 7 Then: Exit Sub
Else
If MsgBox("Everybody to turn off the items?", vbQuestion + vbYesNo, "Atenção!!") = 7 Then: Exit Sub

End If
lstVariaveis.Clear
lstCamposMail.Clear
lstEtiqsMail.Clear
ApagaAspMail
Label28 = 0

Case 3
If lstVariaveis.ListCount = 0 Then: Exit Sub
vEdita = lstEtiqsMail.ListIndex + 1

frmEditar.Show vbModal

End Select

Exit Sub

ErroExc:
If Idioma = "BR" Then
If Err = 5 Then: MsgBox "Selecione um campo para excluir!!", vbCritical, "Aviso de Erro!!": Exit Sub
Else
If Err = 5 Then: MsgBox "Select a field to exclude!!", vbCritical, "Aviso de Erro!!": Exit Sub
End If

If Err = 364 Then: Exit Sub
If Err = 3021 Then: GoTo erro3021: Exit Sub
MsgBox Err & " <-> " & Err.Description, vbCritical, "Aviso de Erro!!"
End Sub

Private Sub lblNovo_Click()
novoProjeto
End Sub

Private Sub lblOk_Click(Index As Integer)
ok
End Sub

Private Sub lblPaginar_Click(Index As Integer)
Select Case Index
Dim XcOR As Variant
Dim xVD As Variant, xAZ As Variant, xVM As Variant

Case 0
    If txtPaginar.Text = 0 Then: Exit Sub
    txtPaginar = txtPaginar - 1
    gravaMudancas
Case 1
    txtPaginar = txtPaginar + 1
    gravaMudancas
Case 2
    txtTamCampo = txtTamCampo + 1
    gravaMudancas
Case 3
    If txtTamCampo.Text = 0 Then: Exit Sub
    txtTamCampo = txtTamCampo - 1
    gravaMudancas
Case 4
    If txtQuant.Text = 0 Then: Exit Sub
    txtQuant = txtQuant - 1
    gravaMudancas
Case 5
    txtQuant = txtQuant + 1
    gravaMudancas
End Select

End Sub

Private Sub lblRegistro_Click()
registrar
End Sub

Private Sub lblSair_Click()
prSair
End Sub

Private Sub lblTAlteracao_Click()
txtCriaConexao.Visible = False
txtCriaInclusao.Visible = False
txtCriaConsulta.Visible = False
txtCriaAlteracao.Visible = True
End Sub

Private Sub lblTConexao_Click()
txtCriaConexao.Visible = True
txtCriaInclusao.Visible = False
txtCriaConsulta.Visible = False
txtCriaAlteracao.Visible = False
End Sub

Private Sub lblTConsulta_Click()
txtCriaConexao.Visible = False
txtCriaInclusao.Visible = False
txtCriaConsulta.Visible = True
txtCriaAlteracao.Visible = False
End Sub

Private Sub lblTInclusao_Click()
txtCriaConexao.Visible = False
txtCriaInclusao.Visible = True
txtCriaConsulta.Visible = False
txtCriaAlteracao.Visible = False
End Sub

Private Sub lstCampos_Click()
On Error Resume Next
    Dim vSenha As String, vNulo As String
    xAbrir = 1
        
    OpenDB "Select * from detalhes Where d_nome = '" & lstCampos.Text & "'"
    
    If Not TbDet.BOF And Not TbDet.EOF Then
    txtRot.Text = Format(TbDet("d_rotulo"), "@")
    txtInicial.Text = Format(TbDet("d_valorInicial"), "@")
    txtTamCampo.Text = TbDet("d_tamaho")
    txtQuant.Text = TbDet("d_maximo")
    
    vSenha = TbDet("d_senha").Value
    If vSenha = "S" Then
    optSenha(0).Value = True
    optSenha(1).Value = False
    Else
    optSenha(1).Value = True
    optSenha(0).Value = False
    End If
    
    If TbDet("d_exibir") = "S" Then
        ckExibir.Value = 1
    Else
        ckExibir.Value = 0
    End If
    
    cmbTipoCampo.Text = Format(TbDet("d_tipo"), "@")
    
    vNulo = TbDet("d_nulo")
    If vNulo = "S" Then
    ckNulo.Value = 1
    Else
    ckNulo.Value = 0
    End If
    
    End If
    
    If cmbTipoCampo.Text = "Caixa de Texto" Then
    Label11.Caption = "Lagura em Caracteres:"
    Label10.Caption = "Número de Linhas:"
    Text6.Visible = False
    Text5.Visible = True

    Else
    Label11.Caption = "Tamanho deste campo:"
    Label10.Caption = "Quantidade máxima de caracteres:"
    Text6.Visible = True
    Text5.Visible = False

    End If

    CloseDB
    xAbrir = 0

End Sub

Private Sub lstEtiqsMail_Click()
vEdita = lstEtiqsMail.ListIndex + 1
End Sub

Private Sub mnuarq_Click(Index As Integer)
Select Case Index
Case 0
novoProjeto
Case 2
prSair
End Select
End Sub

Private Sub mnuConf_Click(Index As Integer)
Select Case Index
Case 0
frmCaminho.Show vbModal
End Select
End Sub

Private Sub mnucont_Click(Index As Integer)
 Dim sucesso As Integer
 Dim site As String
Select Case Index
Case 0
site = App.Path & "\ajuda\ajuda.htm"
sucesso = ShellToBrowser(Me, site, 0)

Case 1
Label26_Click

Case 2
    site = "http://www.pagemaster.tk"
    sucesso = ShellToBrowser(Me, site, 0)
   
Case 4
   site = "mailto:" & Trim("") & "?Subject=Construa páginas ASP, ASPMail e CDONTS facilmente!!!"
   successo = ShellToBrowser(Me, site, 0)

End Select
End Sub

Private Sub mnuId_Click(Index As Integer)
Select Case Index
Case 0
mnuId(0).Checked = True
mnuId(1).Checked = False
    WritePrivateProfileString _
        "Config", "Idioma", _
        "BR", App.Path & "\pMaster.ini"
        Idioma = "BR"
        CarregaStrings
Case 1
mnuId(0).Checked = False
mnuId(1).Checked = True
    WritePrivateProfileString _
        "Config", "Idioma", _
        "US", App.Path & "\pMaster.ini"
        Idioma = "US"
        CarregaStrings
End Select

End Sub

Private Sub optSenha_Click(Index As Integer)
Select Case Index
Case 0
If cmbTipoCampo.Text = "Caixa de Texto" Then
optSenha(1).Value = 1
Exit Sub
Else
Text6.PasswordChar = "*"
End If
Case 1
Text6.PasswordChar = ""

End Select
End Sub

Private Sub optSenha_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
gravaMudancas
End Sub

Private Sub picAbrir_Click()
AbrirProcura
End Sub

Private Sub picConfig_Click()
ConfigASP
End Sub

Private Sub picConfigMail_Click()
pcasp
End Sub

Private Sub picCriaPaginaOk_Click()
freGeradoAsp.Visible = False
txtGerado.Text = ""
txtGeradoP2.Text = ""
End Sub

Private Sub picDefineMail_Click()
ldasp
End Sub

Private Sub picDefinir_Click()
definir
End Sub

Private Sub picEnviarDados_Click()
txtGerado.Visible = False
txtGeradoP2.Visible = True
End Sub

Private Sub picGeraMail_Click()
CriarAspMail
End Sub

Private Sub picGerar_Click()
CriarConexao
End Sub

Private Sub picInserirDados_Click()
txtGerado.Visible = True
txtGeradoP2.Visible = False
End Sub

Private Sub picMail_Click(Index As Integer)
On Error GoTo ErroExc
Select Case Index
Case 0
If Idioma = "BR" Then
If txtCampo.Text = "" Then: MsgBox "Informe o nome do campo!", vbInformation: txtCampo.SetFocus: Exit Sub
Else
If txtCampo.Text = "" Then: MsgBox "Inform the name of the field!", vbInformation: txtCampo.SetFocus: Exit Sub
End If

'-------------Procura por nome duplicado------------
    OpenDB "Select am_descricao From aspmail Where am_descricao='" & txtCampo.Text & ":'"
    
    If TbDet("am_descricao") = txtCampo.Text & ":" Then
    If Err = 3021 Then: GoTo erro3021
    
    If Idioma = "BR" Then
    MsgBox "Nome de Campo duplicado!", vbCritical, "Atenção!!"
    Else
    MsgBox "Name of duplicated Field!", vbCritical, "Attention!!"
    End If
    
    CloseDB
    
    Exit Sub
    End If
    
erro3021:
    CloseDB

'----------------------------------------------------
lstEtiqsMail.AddItem txtCampo.Text & ":"
Label28 = Label28 + 1
Dim p As String, i As String, w As String, Y As String, X As String
vCampo = Replace(txtCampo, " ", "", 1)
lstVariaveis.AddItem "v" & lstEtiqsMail.ListCount & Left(UCase(vCampo), 1) & Left(LCase(vCampo), 15)
lstCamposMail.AddItem "t" & lstEtiqsMail.ListCount & Left(UCase(vCampo), 1) & Left(LCase(vCampo), 15)


'------------------------------------ gravar
    OpenDB "aspmail"
    
    TbDet.AddNew
    TbDet("am_codigo") = lstEtiqsMail.ListCount
    TbDet("am_descricao") = txtCampo.Text & ":"
    
    TbDet("am_variavel") = "v" & lstEtiqsMail.ListCount & Left(UCase(vCampo), 1) & Left(LCase(vCampo), 15)
    TbDet("am_campo") = "t" & lstEtiqsMail.ListCount & Left(UCase(vCampo), 1) & Left(LCase(vCampo), 15)

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

'------------------------------fim da gravação

ckObrigatorio.Value = 0
ckSenha.Value = 0
txtCampo.Text = ""
txtCampo.SetFocus

Case 1
If lstVariaveis.ListCount = 0 Then: Exit Sub

lstVariaveis.RemoveItem lstEtiqsMail.ListIndex
lstCamposMail.RemoveItem lstEtiqsMail.ListIndex
lstEtiqsMail.RemoveItem lstEtiqsMail.ListIndex
Label28 = Label28 - 1

Case 2
If lstVariaveis.ListCount = 0 Then: Exit Sub
If Idioma = "BR" Then
If MsgBox("Deseja apagar todos os itens?", vbQuestion + vbYesNo, "Atenção!!") = 7 Then: Exit Sub
Else
If MsgBox("Everybody to turn off the items?", vbQuestion + vbYesNo, "Atenção!!") = 7 Then: Exit Sub

End If
lstVariaveis.Clear
lstCamposMail.Clear
lstEtiqsMail.Clear
ApagaAspMail
Label28 = 0

Case 3
If lstVariaveis.ListCount = 0 Then: Exit Sub
vEdita = lstEtiqsMail.ListIndex + 1

frmEditar.Show vbModal

End Select

Exit Sub

ErroExc:
If Idioma = "BR" Then
If Err = 5 Then: MsgBox "Selecione um campo para excluir!!", vbCritical, "Aviso de Erro!!": Exit Sub
Else
If Err = 5 Then: MsgBox "Select a field to exclude!!", vbCritical, "Aviso de Erro!!": Exit Sub
End If

If Err = 364 Then: Exit Sub
If Err = 3021 Then: GoTo erro3021: Exit Sub
MsgBox Err & " <-> " & Err.Description, vbCritical, "Aviso de Erro!!"

End Sub

Private Sub picNovo_Click()
novoProjeto
End Sub

Private Sub picOk_Click(Index As Integer)
ok
End Sub

Private Sub picRegistro_Click()
registrar
End Sub

Private Sub picSair_Click()
prSair
End Sub

Public Sub abrirFreme()
picNovo.Visible = False
picGeral.Visible = False
freConfig.Visible = False
freCampos.Visible = False
freConfCampo.Visible = False
freOpt.Visible = True
freOpt.Top = 2535
freOpt.Left = 1590
End Sub

Public Sub fecharFreme()
picNovo.Visible = True
freOpt.Visible = False
End Sub

Public Sub ConfigASP()
freConfig.Top = 945
freConfig.Left = 315
freConfig.Visible = True
freCampos.Visible = False
freConfCampo.Visible = False
freForms.Visible = False
End Sub

Public Sub ok()
If optAsp.Value = True Then
    fecharFreme
    AbrirProcura
    picGeral.Top = 45
    picGeral.Left = 1695
    picGeral.Visible = True
    freConfCampo.Top = 750
    freConfCampo.Left = 2550
    freCampos.Top = 750
    freCampos.Left = 60
    freCampos.Visible = True
    freConfCampo.Visible = True
    picAbrir.Visible = False
ElseIf optMail.Value = True Then
    fecharFreme
    freAspMail.Top = 840
    freAspMail.Left = 240
    If Idioma = "BR" Then
    freAspMail.Caption = LoadResString(148)
    Else
    freAspMail.Caption = LoadResString(216)
    End If
    txtServidor.Visible = True
    Label27.Visible = True
    freAspMail.Visible = True
    picGeralMail.Top = 45
    picGeralMail.Left = 1695
    picGeralMail.Visible = True
ElseIf optCdonts.Value = True Then
    fecharFreme
    freAspMail.Top = 840
    freAspMail.Left = 240
    txtServidor.Text = "Null"
    txtServidor.Visible = False
    Label27.Visible = False
    If Idioma = "BR" Then
    freAspMail.Caption = LoadResString(149)
    Else
    freAspMail.Caption = LoadResString(217)
    End If
    freAspMail.Visible = True
    picGeralMail.Top = 45
    picGeralMail.Left = 1695
    picGeralMail.Visible = True
End If

End Sub

Public Sub definir()
freConfig.Visible = False
freForms.Visible = False
freConfCampo.Top = 750
freConfCampo.Left = 2550
freCampos.Top = 750
freCampos.Left = 60
freCampos.Visible = True
freConfCampo.Visible = True

End Sub

Public Sub gravaMudancas()
On Error Resume Next
If xAbrir = 1 Then: Exit Sub
        
    OpenDB "Select * from detalhes Where d_nome = '" & lstCampos.Text & "'"
'//---------Mostra itens da tabela
    TbDet.Edit
    TbDet("d_rotulo") = txtRot.Text
    TbDet("d_valorInicial") = txtInicial.Text
    TbDet("d_tamaho") = txtTamCampo.Text
    TbDet("d_maximo") = txtQuant.Text
    TbDet("d_tipo") = cmbTipoCampo.Text
    
    If optSenha(0).Value = True Then
        TbDet("d_senha") = "S"
    Else
        TbDet("d_senha") = "N"
    End If
    
    If ckNulo.Value = 1 Then
        TbDet("d_nulo") = "S"
    Else
        TbDet("d_nulo") = "N"
    End If
    
    If ckExibir.Value = 1 Then
        TbDet("d_exibir") = "S"
    Else
        TbDet("d_exibir") = "N"
    End If
    
    TbDet.Update
    
    CloseDB
'//-----------Fim da inclusão

End Sub

Private Sub Picture2_Click()

End Sub

Private Sub picTAlteracao_Click()
txtCriaConexao.Visible = False
txtCriaInclusao.Visible = False
txtCriaConsulta.Visible = False
txtCriaAlteracao.Visible = True
End Sub

Private Sub picTConexao_Click()
txtCriaConexao.Visible = True
txtCriaInclusao.Visible = False
txtCriaConsulta.Visible = False
txtCriaAlteracao.Visible = False
End Sub

Private Sub picTConsulta_Click()
txtCriaConexao.Visible = False
txtCriaInclusao.Visible = False
txtCriaConsulta.Visible = True
txtCriaAlteracao.Visible = False
End Sub

Private Sub picTInclusao_Click()
txtCriaConexao.Visible = False
txtCriaInclusao.Visible = True
txtCriaConsulta.Visible = False
txtCriaAlteracao.Visible = False
End Sub

Private Sub Picture1_Click()
obter
End Sub

Private Sub Timer1_Timer()
txtCampo.SetFocus
Timer1.Interval = 0
End Sub

Private Sub txtInicial_KeyUp(KeyCode As Integer, Shift As Integer)
gravaMudancas
End Sub

Public Sub apagarBanco()
On Error Resume Next
abrirBanco

Do While Not TbDet.EOF
TbDet.Delete
TbDet.MoveNext
Loop

    TbDet.Close
    Banco.Close
    Set TbDet = Nothing
    Set Banco = Nothing

End Sub

Private Sub txtQuant_KeyUp(KeyCode As Integer, Shift As Integer)
gravaMudancas
End Sub

Private Sub txtRot_KeyUp(KeyCode As Integer, Shift As Integer)
gravaMudancas
End Sub

Private Sub txtTamCampo_KeyUp(KeyCode As Integer, Shift As Integer)
gravaMudancas
End Sub

Public Sub novoProjeto()
If Idioma = "BR" Then
If MsgBox(LoadResString(178), vbQuestion + vbYesNo, "Atenção!!") = 7 Then: Exit Sub
Else
If MsgBox(LoadResString(278), vbQuestion + vbYesNo, "Attention!!") = 7 Then: Exit Sub
End If

picAbrir.Visible = True
picGeralMail.Visible = False
apagarBanco
ApagaAspMail
abrirFreme
freAspMail.Visible = False
freConfigAspMail.Visible = False
freGeradoAsp.Visible = False
freForms.Visible = False

txtServidor.Text = ""
txtAssunto.Text = ""
txtRementente.Text = ""
txtNomeRemet.Text = ""
txtDestino.Text = ""
txtTituloPagina.Text = ""
txtNomeArq.Text = ""
txtNomeArquivo.Text = ""

lstVariaveis.Clear
lstCamposMail.Clear
lstEtiqsMail.Clear

End Sub


Public Sub ApagaAspMail()
On Error Resume Next
    NomeArq = "\campos.mdb"
    PathApp = App.Path & NomeArq
    NomeTabela = "aspmail"
    Set Banco = OpenDatabase(PathApp)
    Set TbDet = Banco.OpenRecordset(NomeTabela, dbOpenDynaset)

Do While Not TbDet.EOF
TbDet.Delete
TbDet.MoveNext
Loop

    TbDet.Close
    Banco.Close
    Set TbDet = Nothing
    Set Banco = Nothing

End Sub

Public Sub pcasp()
freAspMail.Visible = False
freGeradoAsp.Visible = False
freConfigAspMail.Top = 945
freConfigAspMail.Left = 315

If Idioma = "BR" Then
    If optMail.Value = True Then
        freConfigAspMail.Caption = LoadResString(127)
    Else
        freConfigAspMail.Caption = LoadResString(263)
    End If
Else
    If optMail.Value = True Then
        freConfigAspMail.Caption = LoadResString(238)
    Else
        freConfigAspMail.Caption = LoadResString(264)
    End If
End If

freConfigAspMail.Visible = True

End Sub

Public Sub ldasp()
freConfigAspMail.Visible = False
freGeradoAsp.Visible = False
freAspMail.Visible = True
freAspMail.Top = 840
freAspMail.Left = 240

End Sub

Public Sub CriarAspMail()
If Idioma = "BR" Then
If Label28.Caption = 0 Then: MsgBox LoadResString(179), vbInformation, "Atenção!!": Exit Sub

If txtServidor.Text = "" Then: MsgBox LoadResString(180), vbInformation, "Atenção!!": _
pcasp: txtServidor.SetFocus: Exit Sub

If txtAssunto.Text = "" Then: MsgBox LoadResString(181), vbInformation, "Atenção!!": _
pcasp: txtAssunto.SetFocus: Exit Sub

If txtRementente.Text = "" Then: MsgBox LoadResString(182), vbInformation, "Atenção!!": _
pcasp: txtRementente.SetFocus: Exit Sub

If txtNomeRemet.Text = "" Then: MsgBox LoadResString(183), vbInformation, "Atenção!!": _
pcasp: txtNomeRemet.SetFocus: Exit Sub

If txtDestino.Text = "" Then: MsgBox LoadResString(274), vbInformation, "Atenção!!": _
pcasp: txtDestino.SetFocus: Exit Sub

If txtTituloPagina.Text = "" Then: MsgBox LoadResString(184), vbInformation, "Atenção!!": _
pcasp: txtTituloPagina.SetFocus: Exit Sub

If txtNomeArq.Text = "" Then: MsgBox LoadResString(185), vbInformation, "Atenção!!": _
pcasp: txtNomeArq.SetFocus: Exit Sub

Else
If Label28.Caption = 0 Then: MsgBox LoadResString(266), vbInformation, LoadResString(273): Exit Sub

If txtServidor.Text = "" Then: MsgBox LoadResString(267), vbInformation, LoadResString(273): _
pcasp: txtServidor.SetFocus: Exit Sub

If txtAssunto.Text = "" Then: MsgBox LoadResString(268), vbInformation, LoadResString(273): _
pcasp: txtAssunto.SetFocus: Exit Sub

If txtRementente.Text = "" Then: MsgBox LoadResString(269), vbInformation, LoadResString(273): _
pcasp: txtRementente.SetFocus: Exit Sub

If txtNomeRemet.Text = "" Then: MsgBox LoadResString(270), vbInformation, LoadResString(273): _
pcasp: txtNomeRemet.SetFocus: Exit Sub

If txtDestino.Text = "" Then: MsgBox LoadResString(275), vbInformation, LoadResString(273): _
pcasp: txtDestino.SetFocus: Exit Sub

If txtTituloPagina.Text = "" Then: MsgBox LoadResString(271), vbInformation, LoadResString(273): _
pcasp: txtTituloPagina.SetFocus: Exit Sub

If txtNomeArq.Text = "" Then: MsgBox LoadResString(272), vbInformation, LoadResString(273): _
pcasp: txtNomeArq.SetFocus: Exit Sub
End If
freConfigAspMail.Visible = False

length = GetPrivateProfileString( _
"Config", "Salvar", App.Path, _
buf, Len(buf), App.Path & "\pMaster.ini")
Caminho = Left$(buf, length) & "\"

If Dir$(Caminho & txtNomeArq.Text & ".asp") <> "" Then

    If Idioma = "BR" Then
If MsgBox("O arquivo " & Caminho & txtNomeArq & ".asp" & Chr(10) & _
"já existe. Deseja substitui-lo?", vbQuestion + vbYesNo) = vbNo Then: Exit Sub
    Else
If MsgBox("The file " & Caminho & txtNomeArq & ".asp" & Chr(10) & _
"it already exists. Wants to substitute it?", vbQuestion + vbYesNo) = vbNo Then: Exit Sub
    End If
End If

vTexto = vTexto & "<html>" & vbCrLf
vTexto = vTexto & "<head>" & vbCrLf
vTexto = vTexto & "<title>" & txtTituloPagina & "</title>" & vbCrLf
vTexto = vTexto & "</head>" & vbCrLf
vTexto = vTexto & "<body>" & vbCrLf
vTexto = vTexto & "<center><font face=Arial size=4><b>Formulário AspMail</b></font></h2></center>" & vbCrLf

vTexto = vTexto & "<script Language=" & Chr(34) & "JavaScript" & Chr(34) & ">" & vbCrLf
vTexto = vTexto & "function Validator(theForm)" & vbCrLf
vTexto = vTexto & "{" & vbCrLf

    sql = "Select * From AspMail Where am_obrigatorio = 'S'"
    abreBase
    Do While Not Tb.EOF
vTexto = vTexto & "  if (theForm." & Tb("am_campo") & ".value == " & Chr(34) & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "  {" & vbCrLf
vTexto = vTexto & "    alert(" & Chr(34) & "Digite um valor para o campo " & Tb("am_descricao") & Chr(34) & ");" & vbCrLf
vTexto = vTexto & "    theForm." & Tb("am_campo") & ".focus();" & vbCrLf
vTexto = vTexto & "    return (false);" & vbCrLf
vTexto = vTexto & "  }" & vbCrLf
    Tb.MoveNext
    Loop
    Tb.Close
    
vTexto = vTexto & "  return (true);" & vbCrLf
vTexto = vTexto & "}" & vbCrLf
vTexto = vTexto & "</script>" & vbCrLf
    
    Dim vTipoManda As String
    If Option1.Value = True Then
        vTipoManda = "GET"
    Else
        vTipoManda = "POST"
    End If
    
vTexto = vTexto & "        <form method=" & vTipoManda & " name=Manda onsubmit=" & Chr(34) & "return Validator(this)" & Chr(34) & " Action=" & Chr(34) & txtNomeArq & "E.asp" & Chr(34) & ">" & vbCrLf
vTexto = vTexto & "          <table border=0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
vTexto = vTexto & "            <tr>" & vbCrLf

    sql = "Select * From AspMail"
    abreBase
    Do While Not Tb.EOF
vTexto = vTexto & "              <td width=" & Chr(34) & "28%" & Chr(34) & "><p align=right style=" & Chr(34) & "margin-right: 5" & Chr(34) & "><font color=#000000 face=Verdana size=2><b>" & Tb("am_descricao") & "</b></font></td>" & vbCrLf
vTexto = vTexto & "              <center>" & vbCrLf
    If Tb("am_senha") = "N" Then
vTexto = vTexto & "              <td width=" & Chr(34) & "72%" & Chr(34) & "><input type=text name=" & Tb("am_campo") & " size=52></td>" & vbCrLf
    Else
vTexto = vTexto & "              <td width=" & Chr(34) & "72%" & Chr(34) & "><input type=password name=" & Tb("am_campo") & " size=52></td>" & vbCrLf
    End If
vTexto = vTexto & "              </tr>" & vbCrLf
    Tb.MoveNext
    Loop
    Tb.Close
    If Check3.Value = 1 Then
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "              <tr>" & vbCrLf
vTexto = vTexto & "            <td width=28% valign=top><p align=right style=" & Chr(34) & "margin-right: 5" & Chr(34) & "><font color=#000000 face=Verdana size=2><b>MENSAGEM:</b></font></td>" & vbCrLf
vTexto = vTexto & "            <center>" & vbCrLf
vTexto = vTexto & "            <td width=72><textarea rows=6 name=txtMensagem cols=41></textarea></td>" & vbCrLf
vTexto = vTexto & "              </tr>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
    End If
vTexto = vTexto & "             <tr>" & vbCrLf
vTexto = vTexto & "         <td width=760 colspan=2>" & vbCrLf
vTexto = vTexto & "          <p align=center><input type=submit value=Enviar name=cmdEnvio></p>" & vbCrLf
vTexto = vTexto & "         </td>" & vbCrLf
vTexto = vTexto & "             </tr>" & vbCrLf
vTexto = vTexto & "            </center>" & vbCrLf
vTexto = vTexto & "              <center>" & vbCrLf
vTexto = vTexto & "            </table>" & vbCrLf
vTexto = vTexto & "          </form>" & vbCrLf
vTexto = vTexto & "</body>"

txtGerado.Text = vTexto
SaveFileAs Caminho & txtNomeArq & ".asp", txtGerado
vTexto = ""
CriarAspMailP2
End Sub

Public Sub abreBase()
    NomeArq = "\campos.mdb"
    PathApp = App.Path & NomeArq
    Set Banco = OpenDatabase(PathApp)
    Set Tb = Banco.OpenRecordset(sql, dbOpenDynaset)
End Sub

Sub SaveFileAs(Filename As String, Caixa As String)
    On Error Resume Next
    Dim strContents As String

    Open Filename For Output As #1
    strContents = Caixa
    Screen.MousePointer = 11
    Print #1, strContents
    Close #1
    Screen.MousePointer = 0
    If Err Then
        MsgBox Error, 48, App.Title
    End If
End Sub

Public Sub CriarAspMailP2()

vTexto = vTexto & "<%" & vbCrLf
vTexto = vTexto & "html = " & Chr(34) & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "<html>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "<head>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "<title>" & txtTituloPagina & "</title>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "</head>" & Chr(34) & vbCrLf
vTexto = vTexto & "Response.Write html" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->" & vbCrLf
    If Idioma = "BR" Then
vTexto = vTexto & "'<------ /  Gerando com PageMaster   \ ------>" & vbCrLf
    Else
vTexto = vTexto & "'<------ / Generating with PageMaster\ ------>" & vbCrLf
    End If
vTexto = vTexto & "'<----- /  http:\\www.pagemaster.tk   \ ----->" & vbCrLf
vTexto = vTexto & "'<----- \_____________________________/ ----->" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "    Dim sMsgErr" & vbCrLf
vTexto = vTexto & "" & vbCrLf
'    If Idioma = "BR" Then
'vTexto = vTexto & "    '***** Executa as ações desta página *****" & vbCrLf
'    Else
'vTexto = vTexto & "    '***** It executes the actions of this page *****" & vbCrLf
'    End If
vTexto = vTexto & "    Sub ProcessaPagina()" & vbCrLf
'    If Idioma = "BR" Then
'vTexto = vTexto & "    '***** Declaração das Variáveis *****" & vbCrLf
'    Else
'vTexto = vTexto & "    '***** Declaration of the Variables *****" & vbCrLf
'    End If
    sql = "Select * From AspMail"
    abreBase
    Do While Not Tb.EOF
vTexto = vTexto & "    Dim " & Tb("am_variavel") & vbCrLf
    Tb.MoveNext
    Loop
    Tb.Close
    
    If Check3.Value = 1 Then
vTexto = vTexto & "    Dim vMmensagem" & vbCrLf
    End If

'    If Idioma = "BR" Then
vTexto = vTexto & "    sEMail = " & Chr(34) & LCase(txtDestino.Text) & Chr(34) & vbCrLf
'    Else
'vTexto = vTexto & "    sEMail = " & Chr(34) & LCase(txtDestino.Text) & Chr(34) & " 'when I want that the message comes for me" & vbCrLf
'    End If
    
vTexto = vTexto & "" & vbCrLf
'    If Idioma = "BR" Then
'vTexto = vTexto & "    '***** Obtém valores preenchidos no Formulário *****" & vbCrLf
'    Else
'vTexto = vTexto & "    '***** Obtains values filled in the Form *****" & vbCrLf
'    End If
    
    Dim vTipoManda As String
    If Option1.Value = True Then
        vTipoManda = "QueryString"
    Else
        vTipoManda = "Form"
    End If

    abreBase
    Do While Not Tb.EOF
vTexto = vTexto & "    " & Tb("am_variavel") & " = Request." & vTipoManda & "(" & Chr(34) & Tb("am_campo") & Chr(34) & ")" & vbCrLf
    Tb.MoveNext
    Loop
    Tb.Close
    
    If Check3.Value = 1 Then
vTexto = vTexto & "    vMmensagem = Request." & vTipoManda & "(" & Chr(34) & "txtMENSAGEM" & Chr(34) & ")" & vbCrLf
    End If

vTexto = vTexto & "" & vbCrLf
'    If Idioma = "BR" Then
'vTexto = vTexto & "    '***** Monta corpo da mensagem a enviar por e-mail" & vbCrLf
'    Else
'vTexto = vTexto & "    '***** It sets up body of the message to send for e-mail" & vbCrLf
'    End If
    
    abreBase
    Do While Not Tb.EOF
vTexto = vTexto & "    sBodyText = sBodyText & " & Chr(34) & Tb("am_descricao") & " " & Chr(34) & " & " & Tb("am_variavel") & " & " & "vbCrLf   'corpo" & vbCrLf
    Tb.MoveNext
    Loop
    Tb.Close
    
    If Check3.Value = 1 Then
vTexto = vTexto & "    sBodyText = sBodyText & " & Chr(34) & "MENSAGEM:" & " " & Chr(34) & " & vMmensagem & " & "vbCrLf   'corpo" & vbCrLf
    End If
    
vTexto = vTexto & "    sBodyText = sBodyText & " & Chr(34) & "----------------------------------------------" & Chr(34) & " & vbCrLf 'corpo" & vbCrLf
vTexto = vTexto & "" & vbCrLf
'    If Idioma = "BR" Then
'vTexto = vTexto & "    '***** Envia E-Mail para o destinatário *****" & vbCrLf
'    Else
'vTexto = vTexto & "    '***** Sends E-mail for the addressee *****" & vbCrLf
'    End If
vTexto = vTexto & "    On Error Resume Next" & vbCrLf
    
    If optMail.Value = True Then
vTexto = vTexto & "    Set Mail = Server.CreateObject(" & Chr(34) & "Persits.MailSender" & Chr(34) & ")" & vbCrLf
    
    If Idioma = "BR" Then
vTexto = vTexto & "    Mail.Host = " & Chr(34) & LCase(txtServidor.Text) & Chr(34) & " ' Especifique o nome do seu servidor SMTP." & vbCrLf
vTexto = vTexto & "    Mail.From = " & Chr(34) & LCase(txtRementente.Text) & Chr(34) & " ' Remetente da mensagem" & vbCrLf
vTexto = vTexto & "    Mail.FromName = " & Chr(34) & txtNomeRemet.Text & Chr(34) & " ' Nome do remetente" & vbCrLf
vTexto = vTexto & "    Mail.AddAddress sEMail ' Destinatario da mensagem para" & vbCrLf
vTexto = vTexto & "    Mail.Subject = " & Chr(34) & txtAssunto.Text & Chr(34) & " 'assunto" & vbCrLf
vTexto = vTexto & "    Mail.Body = sBodyText 'corpo da mensagem montada acima" & vbCrLf
    Else
vTexto = vTexto & "    Mail.Host = " & Chr(34) & LCase(txtServidor.Text) & Chr(34) & " ' Specify server SMTP name." & vbCrLf
vTexto = vTexto & "    Mail.From = " & Chr(34) & LCase(txtRementente.Text) & Chr(34) & " ' Remittent of the message" & vbCrLf
vTexto = vTexto & "    Mail.FromName = " & Chr(34) & txtNomeRemet.Text & Chr(34) & " ' Name of the remittent" & vbCrLf
vTexto = vTexto & "    Mail.AddAddress sEMail ' Addressee of the message" & vbCrLf
vTexto = vTexto & "    Mail.Subject = " & Chr(34) & txtAssunto.Text & Chr(34) & " 'subject" & vbCrLf
vTexto = vTexto & "    Mail.Body = sBodyText 'body of the mounted message above" & vbCrLf
    End If
    
    Else
vTexto = vTexto & "    Set Mail = Server.CreateObject(" & Chr(34) & "CDONTS.NewMail" & Chr(34) & ")" & vbCrLf
'vTexto = vTexto & "    Mail.Host = " & Chr(34) & LCase(txtServidor.Text) & Chr(34) & " ' Especifique o nome do seu servidor SMTP." & vbCrLf
vTexto = vTexto & "    Mail.From = " & Chr(34) & LCase(txtRementente.Text) & Chr(34) & " ' Remetente da mensagem" & vbCrLf
'vTexto = vTexto & "    Mail.FromName = " & Chr(34) & txtNomeRemet.Text & Chr(34) & " ' Nome do remetente" & vbCrLf
vTexto = vTexto & "    Mail.To = sEMail ' Destinatario da mensagem para" & vbCrLf
vTexto = vTexto & "    Mail.Subject = " & Chr(34) & txtAssunto.Text & Chr(34) & " 'assunto" & vbCrLf
vTexto = vTexto & "    Mail.Body = sBodyText 'corpo da mensagem montada acima" & vbCrLf
    End If
vTexto = vTexto & "    sMsgErr = " & Chr(34) & Chr(34) & vbCrLf
vTexto = vTexto & "    Mail.Send" & vbCrLf
vTexto = vTexto & "    If Err <> 0 Then" & vbCrLf
    If Idioma = "BR" Then
vTexto = vTexto & "    sMsgErr = " & Chr(34) & "Ocorreu o seguinte erro ao tentar enviar o e-mail: " & Chr(34) & " & Err.Description" & vbCrLf
    Else
vTexto = vTexto & "    sMsgErr = " & Chr(34) & "It happened the following mistake when trying to send the e-mail: " & Chr(34) & " & Err.Description" & vbCrLf
    End If
vTexto = vTexto & "    End If" & vbCrLf
vTexto = vTexto & "    On Error GoTo 0" & vbCrLf
vTexto = vTexto & "    End Sub" & vbCrLf
vTexto = vTexto & "" & vbCrLf
'    If Idioma = "BR" Then
'vTexto = vTexto & "    '***** Executa as ações desta página *****" & vbCrLf
'    Else
'vTexto = vTexto & "    '***** It executes the actions of this page *****" & vbCrLf
'    End If
    
vTexto = vTexto & "    ProcessaPagina" & vbCrLf

vTexto = vTexto & "html = " & Chr(34) & Chr(34) & vbCrLf
vTexto = vTexto & "If (sMsgErr <> " & Chr(34) & Chr(34) & ") Then" & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "<body>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "<div>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "  <table width=100% border=0 cellspacing=0 cellpadding=0 height=21>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "    <tr>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "      <td height=23>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "        <p><font size=3 face=Arial color=#000000>" & Chr(34) & " & sMsgErr & " & Chr(34) & "<br>" & Chr(34) & vbCrLf

vTexto = vTexto & "Else" & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "</font>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "    </tr>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "    <tr>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "      <td height=101>" & Chr(34) & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "    </tr>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "    <tr>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "      <td height=101>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "        <p align=center><font face=Verdana size=2 color=#008080>" & Chr(34) & vbCrLf
    If Idioma = "BR" Then
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "          Sua mensagem foi encaminhada com sucesso.<br>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "          Logo entraremos em contato. </font> <font face=Verdana size=2 color=#800000><br>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "          Agradecemos por seu interesse!<br>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "          <a href=" & Chr(34) & " & chr(34) & " & Chr(34) & "javascript:close()" & Chr(34) & " & chr(34) & " & Chr(34) & ">Fechar</a><br>" & Chr(34) & vbCrLf
    Else
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "          Its message was guided with success.<br>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "          Therefore we will enter in contact. </font> <font face=Verdana size=2 color=#800000><br>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "          We thanked for its interest!<br>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "          <a href=" & Chr(34) & " & chr(34) & " & Chr(34) & "javascript:close()" & Chr(34) & " & chr(34) & " & Chr(34) & ">Close</a><br>" & Chr(34) & vbCrLf
    End If
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "        </font> </p>" & Chr(34) & vbCrLf
vTexto = vTexto & "End If" & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "    </tr>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "  </table>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "</div>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "</body>" & Chr(34) & vbCrLf
vTexto = vTexto & "html = html " & Chr(38) & Chr(32) & Chr(34) & "</html>" & Chr(34) & vbCrLf
vTexto = vTexto & "Response.Write html" & vbCrLf
vTexto = vTexto & "%>"

freGeradoAsp.Top = 705
freGeradoAsp.Left = 240

If optMail.Value = True Then
    freGeradoAsp.Caption = LoadResString(167)
Else
    freGeradoAsp.Caption = LoadResString(168)
End If

freGeradoAsp.Visible = True
txtGeradoP2.Text = vTexto
SaveFileAs Caminho & txtNomeArq.Text & "E.asp", txtGeradoP2
vTexto = ""
If Idioma = "BR" Then
MsgBox "O arquivo foi criado em " & Chr(10) & _
Caminho, vbInformation
Else
MsgBox "The file was created in " & Chr(10) & _
Caminho, vbInformation
End If
End Sub

Public Sub CriarConsulta()
'On Error GoTo errGerar

Dim xLista As String, xBorda As Byte, xCelPart As Byte, vContador As Byte

length = GetPrivateProfileString( _
"Config", "Salvar", App.Path, _
buf, Len(buf), App.Path & "\pMaster.ini")
Caminho = Left$(buf, length) & "\"

If Check2.Value = 1 Then
xBorda = 1
xCelPart = 0
Else
xBorda = 0
xCelPart = 1
End If

vTexto = vTexto & "<%" & vbCrLf
vTexto = vTexto & "'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->" & vbCrLf
vTexto = vTexto & "'<------ /  Gerando com PageMaster   \ ------>" & vbCrLf
vTexto = vTexto & "'<----- /  http:\\www.pagemaster.tk   \ ----->" & vbCrLf
vTexto = vTexto & "'<----- \_____________________________/ ----->" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "'NOME DO ARQUIVO...: " & txtNomeArquivo.Text & ".ASP" & vbCrLf
vTexto = vTexto & "'CRIADO EM.........: " & Now() & vbCrLf
vTexto = vTexto & "'---------------------------------------------" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "Option Explicit" & vbCrLf
vTexto = vTexto & "Const xTmCache = " & txtPaginar.Text & vbCrLf
vTexto = vTexto & "Dim iCount" & vbCrLf
vTexto = vTexto & "Dim sRowColor" & vbCrLf
vTexto = vTexto & "Dim objDB" & vbCrLf
vTexto = vTexto & "Dim objRS" & vbCrLf
vTexto = vTexto & "Dim sDBName" & vbCrLf
vTexto = vTexto & "Dim rsTemp" & vbCrLf
vTexto = vTexto & "Dim Sql" & vbCrLf
vTexto = vTexto & "Dim PagAtual" & vbCrLf
vTexto = vTexto & "Dim TotalPag" & vbCrLf
vTexto = vTexto & "Dim XAltera" & vbCrLf
vTexto = vTexto & "Dim Xcod" & vbCrLf
vTexto = vTexto & "Dim nPags" & vbCrLf

vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "XAltera = Request.QueryString(" & Chr(34) & "XAl" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "Xcod = Request.QueryString(" & Chr(34) & "Cod" & Chr(34) & ")" & vbCrLf

vTexto = vTexto & "Session(" & Chr(34) & "PrimeiraVez" & Chr(34) & ") = request.querystring(" & Chr(34) & "Primeira" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "If Session(" & Chr(34) & "PrimeiraVez" & Chr(34) & ") <> " & Chr(34) & "Nao" & Chr(34) & " Then" & vbCrLf

'Abre a conexao com a base de dados
vTexto = vTexto & "%>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<!--#include file=" & Chr(34) & txtNomeArquivo.Text & "_CNX.asp" & Chr(34) & "-->" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<%" & vbCrLf

vTexto = vTexto & "Set rsTemp = Server.CreateObject(" & Chr(34) & "ADODB.Recordset" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "rsTemp.CacheSize = xTmCache" & vbCrLf
vTexto = vTexto & "rsTemp.PageSize = xTmCache" & vbCrLf
vTexto = vTexto & "" & vbCrLf

vTexto = vTexto & "If XAltera = 1 Then" & vbCrLf
    OpenDB "Select Top 1 d_nome From Detalhes"
vTexto = vTexto & "   objDB.Execute " & Chr(34) & "DELETE FROM " & lblNomeTabela.Caption & " WHERE " & TbDet("d_nome") & " = " & Chr(34) & " & Xcod" & vbCrLf
    CloseDB
vTexto = vTexto & "   SQL = " & Chr(34) & "Select * From " & lblNomeTabela.Caption & Chr(34) & vbCrLf
vTexto = vTexto & "   rsTemp.Open SQL, objDB,3,3" & vbCrLf
vTexto = vTexto & "   Session(" & Chr(34) & "Pagina" & Chr(34) & ") = 1" & vbCrLf
vTexto = vTexto & "   InicioX" & vbCrLf
vTexto = vTexto & "   MostraDados" & vbCrLf

vTexto = vTexto & "Else" & vbCrLf
vTexto = vTexto & "   SQL = " & Chr(34) & "Select * From " & lblNomeTabela.Caption & Chr(34) & vbCrLf
vTexto = vTexto & "   rsTemp.Open SQL, objDB,3,3" & vbCrLf
vTexto = vTexto & "   Session(" & Chr(34) & "Pagina" & Chr(34) & ") = 1" & vbCrLf
vTexto = vTexto & "   InicioX" & vbCrLf
vTexto = vTexto & "   MostraDados" & vbCrLf
vTexto = vTexto & "End If" & vbCrLf

vTexto = vTexto & "" & vbCrLf

vTexto = vTexto & "   Session(" & Chr(34) & "PrimeiraVez" & Chr(34) & ") = " & Chr(34) & "Nao" & Chr(34) & vbCrLf
vTexto = vTexto & "Else" & vbCrLf
vTexto = vTexto & "If Request(" & Chr(34) & "Navegacao" & Chr(34) & ") = " & Chr(34) & "Proxima" & Chr(34) & " Then" & vbCrLf
vTexto = vTexto & "   Session(" & Chr(34) & "Pagina" & Chr(34) & ") = Session(" & Chr(34) & "Pagina" & Chr(34) & ") + 1" & vbCrLf
vTexto = vTexto & "ElseIf Request(" & Chr(34) & "Navegacao" & Chr(34) & ") = " & Chr(34) & "Anterior" & Chr(34) & " Then" & vbCrLf
vTexto = vTexto & "   Session(" & Chr(34) & "Pagina" & Chr(34) & ") = Session(" & Chr(34) & "Pagina" & Chr(34) & ") - 1" & vbCrLf
vTexto = vTexto & "Else" & vbCrLf
vTexto = vTexto & "   Session(" & Chr(34) & "Pagina" & Chr(34) & ") = Trim(Request(" & Chr(34) & "Navegacao" & Chr(34) & "))" & vbCrLf
vTexto = vTexto & "End If" & vbCrLf

vTexto = vTexto & "" & vbCrLf

vTexto = vTexto & "   Set objDB = Server.CreateObject(" & Chr(34) & "ADODB.Connection" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "   objDB.CursorLocation = 3" & vbCrLf

'Abre a conexao com a base de dados
vTexto = vTexto & "%>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<!--#include file=" & Chr(34) & txtNomeArquivo.Text & "_CNX.asp" & Chr(34) & "-->" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<%" & vbCrLf

vTexto = vTexto & "" & vbCrLf

vTexto = vTexto & "   Set rsTemp = Server.CreateObject(" & Chr(34) & "ADODB.Recordset" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "   rsTemp.CacheSize = xTmCache" & vbCrLf
vTexto = vTexto & "   rsTemp.PageSize = xTmCache" & vbCrLf
vTexto = vTexto & "" & vbCrLf

vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "If XAltera = 1 Then" & vbCrLf
    OpenDB "Select Top 1 d_nome From Detalhes"
vTexto = vTexto & "   objDB.Execute " & Chr(34) & "DELETE FROM " & lblNomeTabela.Caption & " WHERE " & TbDet("d_nome") & " = " & Chr(34) & " & Xcod" & vbCrLf
    CloseDB
vTexto = vTexto & "   SQL = " & Chr(34) & "Select * From " & lblNomeTabela.Caption & Chr(34) & vbCrLf
vTexto = vTexto & "   rsTemp.Open SQL, objDB,3,3" & vbCrLf
vTexto = vTexto & "   InicioX" & vbCrLf
vTexto = vTexto & "   MostraDados" & vbCrLf
vTexto = vTexto & "Else" & vbCrLf
vTexto = vTexto & "   SQL = " & Chr(34) & "Select * From " & lblNomeTabela.Caption & Chr(34) & vbCrLf
vTexto = vTexto & "   rsTemp.Open SQL, objDB,3,3" & vbCrLf
vTexto = vTexto & "   InicioX" & vbCrLf
vTexto = vTexto & "   MostraDados" & vbCrLf
vTexto = vTexto & "End If" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "End If" & vbCrLf

vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "Sub InicioX()" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<html>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<head>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<style>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<!--" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "a:link { font-family:Verdana; font-size:9pt; text-decoration:none; }" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "a:visited { font-family:Verdana; font-size:9pt; text-decoration:none; }" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "a:active { font-family:Verdana; font-size:9pt; text-decoration:none; }" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "a:hover { font-family:Verdana; font-size:9pt; text-decoration:underline; }" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "-->" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</style>" & Chr(34) & ") & vbCrLf" & vbCrLf

vTexto = vTexto & "Response.Write (" & Chr(34) & "<script language=JavaScript>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<!--" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "function confirm_delete() {" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "  return confirm (" & Chr(34) & " & Chr(34) & " & Chr(34) & "Você realmente deseja remover este registro do Sistema." & Chr(34) & " & Chr(34) & " & Chr(34) & ")" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "}" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "// -->" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</script>" & Chr(34) & ") & vbCrLf" & vbCrLf

vTexto = vTexto & "Response.Write (" & Chr(34) & "<title>" & txtConsulta.Text & "</title>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<body topmargin=0>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</head>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<body bgcolor=#eee8aa text=black link=#000099 vlink=#000099 alink=#000099>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<table width=750 align=center border=0 cellpadding=0 cellspacing=1>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<tr>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<td width=750 bgcolor=#eee8aa><font face=Verdana size=2 color=#000000><b>Listagem de " & lblNomeTabela.Caption & "</b></font></td></table>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<div align=center>" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<center>" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<table border=0 width=81% cellspacing=0 cellpadding=0 height=32>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<tr>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<td width=25% height=32 align=center bgcolor=#eee8aa><font face=Verdana size=2><a href=" & txtNomeArquivo.Text & "_Inc.asp>Incluir</a></font></td>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<td width=25% height=32 align=center bgcolor=#eee8aa><font face=Verdana size=2><a href=" & txtNomeArquivo.Text & ".asp>Opção 02</a></font></td>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<td width=25% height=32 align=center bgcolor=#eee8aa><font face=Verdana size=2><a href=" & txtNomeArquivo.Text & ".asp>Opção 03</a></font></td>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<td width=25% height=32 align=center bgcolor=#eee8aa><font face=Verdana size=2><a href=" & txtNomeArquivo.Text & ".asp>Opção 04</a></font></td>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</tr>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</table>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</center>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</div>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<hr>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "End Sub" & vbCrLf
vTexto = vTexto & "" & vbCrLf

vTexto = vTexto & "Sub Navega()" & vbCrLf
vTexto = vTexto & "maxcount = CInt(rsTemp.pagecount)" & vbCrLf
vTexto = vTexto & "registros = rsTemp.RecordCount" & vbCrLf
vTexto = vTexto & "howmanyrecs = 0" & vbCrLf
vTexto = vTexto & "rsTemp.AbsolutePage = CDbl(Session(" & Chr(34) & "Pagina" & Chr(34) & "))" & vbCrLf
vTexto = vTexto & "Response.Write " & Chr(34) & "<font face=Arial size=1>Existem <b>" & Chr(34) & " " & Chr(38) & " rsTemp.RecordCount" & " " & Chr(38) & Chr(32) & Chr(34) & "</b> registros na tabela - Mostrando página <b>" & Chr(34) & " " & Chr(38) & " " & "Session(" & Chr(34) & "Pagina" & Chr(34) & ") " & Chr(38) & Chr(32) & Chr(34) & "</b> de <b>" & Chr(34) & " " & Chr(38) & " rsTemp.PageCount " & Chr(38) & " " & Chr(34) & "</b></font>" & Chr(34) & vbCrLf

vTexto = vTexto & "" & vbCrLf

vTexto = vTexto & "If Session(" & Chr(34) & "Pagina" & Chr(34) & ") <> 1 Then" & vbCrLf
vTexto = vTexto & "Response.Write " & Chr(34) & "<a href=" & Chr(34) & Chr(34) & txtNomeArquivo.Text & ".asp?Navegacao=Anterior" & Chr(38) & "primeira=Nao" & Chr(34) & Chr(34) & "><font face=Verdana size=1> - [Anterior]</font></a>" & Chr(34) & vbCrLf
vTexto = vTexto & "End If" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "If Session(" & Chr(34) & "Pagina" & Chr(34) & ") <> rsTemp.PageCount Then" & vbCrLf
vTexto = vTexto & "Response.Write " & Chr(34) & "<a href=" & Chr(34) & Chr(34) & txtNomeArquivo.Text & ".asp?Navegacao=Proxima" & Chr(38) & "primeira=Nao" & Chr(34) & Chr(34) & "><font face=Verdana size=1> - [Proxima]</font></a>" & Chr(34) & vbCrLf
vTexto = vTexto & "Response.Write " & Chr(34) & "<br>" & Chr(34) & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "Dim xNumPgs" & vbCrLf
vTexto = vTexto & "For xNumPgs = 1 To rsTemp.PageCount - 1" & vbCrLf
vTexto = vTexto & "If xNumPgs < 10 Then" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<a href=" & txtNomeArquivo.Text & ".asp?Navegacao=" & Chr(34) & Chr(38) & " xNumPgs " & Chr(38) & Chr(34) & Chr(38) & "primeira=Nao><font face=Arial size=1>0" & Chr(34) & Chr(32) & Chr(38) & " xNumPgs " & Chr(38) & Chr(32) & Chr(34) & " </font></a>" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "Else" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<a href=" & txtNomeArquivo.Text & ".asp?Navegacao=" & Chr(34) & Chr(38) & " xNumPgs " & Chr(38) & Chr(34) & Chr(38) & "primeira=Nao><font face=Arial size=1>" & Chr(34) & Chr(32) & Chr(38) & " xNumPgs " & Chr(38) & Chr(32) & Chr(34) & " </font></a>" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "End If" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "If xNumPgs = maxcount Then Exit For" & vbCrLf
vTexto = vTexto & "Next" & vbCrLf
vTexto = vTexto & "End If" & vbCrLf
vTexto = vTexto & "End Sub" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "Sub MostraDados()" & vbCrLf
vTexto = vTexto & "Dim Contador" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<div Align=center>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<center>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<table align=center border=" & xBorda & " cellpadding=0 cellspacing=" & xCelPart & ">" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<tr bgcolor=silver>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "'-------Cabeçalho da tabela---------" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<td width=6% align=center><font face=" & cmbFonte.Text & " size=" & cmbTamanho.Text & " color=#000000><b></b>OPÇÕES</font></td>" & Chr(34) & ") & vbCrLf" & vbCrLf
    
    OpenDB "Select d_rotulo From detalhes Where d_exibir = 'S'"
    vContador = 0
    Do While Not TbDet.EOF
    If vContador = 6 Then: GoTo vai6
    vContador = vContador + 1
vTexto = vTexto & "Response.Write (" & Chr(34) & "<td width=11% align=center height=25 colspan=3><font face=" & cmbFonte.Text & " size=" & cmbTamanho.Text & " color=#000000><b></b>" & TbDet("d_rotulo") & "</font></td>" & Chr(34) & ") & vbCrLf" & vbCrLf
    TbDet.MoveNext
    Loop
vai6:
    CloseDB
    
vTexto = vTexto & "Response.Write (" & Chr(34) & "</tr>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Navega" & vbCrLf
vTexto = vTexto & "For Contador = 1 To xTmCache" & vbCrLf

    If Check1.Value = 1 Then
vTexto = vTexto & "    iCount = iCount + 1" & vbCrLf
vTexto = vTexto & "    If iCount Mod 2 = 0 Then" & vbCrLf
vTexto = vTexto & "        sRowColor = " & Chr(34) & "#ffdead" & Chr(34) & vbCrLf
vTexto = vTexto & "    Else" & vbCrLf
vTexto = vTexto & "        sRowColor = " & Chr(34) & "#fafad2" & Chr(34) & vbCrLf
vTexto = vTexto & "    End If" & vbCrLf
    Else
vTexto = vTexto & "        sRowColor = " & Chr(34) & "#eee8aa" & Chr(34) & vbCrLf
    End If

    OpenDB "Select Top 1 d_nome From Detalhes"
vTexto = vTexto & "Response.Write (" & Chr(34) & "<tr bgcolor=" & Chr(34) & Chr(38) & " sRowColor " & Chr(38) & Chr(34) & ">" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<td width=3% align=center><a onclick=" & Chr(34) & " & Chr(34) & " & Chr(34) & "return confirm_delete()" & Chr(34) & " & Chr(34) & " & Chr(34) & "& href=" & txtNomeArquivo.Text & ".asp?Cod=" & Chr(34) & Chr(32) & Chr(38) & Chr(32) & "rsTemp(" & Chr(34) & TbDet("d_nome") & Chr(34) & ")" & Chr(32) & Chr(38) & Chr(32) & Chr(34) & Chr(38) & "XAl=1" & "><img border=0 src=images/Delete.gif alt=Excluir width=16 height=18></a><a href=" & txtNomeArquivo.Text & "_ALT.asp?id=" & Chr(34) & Chr(32) & Chr(38) & Chr(32) & "rsTemp(" & Chr(34) & TbDet("d_nome") & Chr(34) & ")" & Chr(32) & Chr(38) & Chr(32) & Chr(34) & "><img border=0 src=images/edit.gif alt=Editar align=absmiddle width=16 height=18></a></td>" & Chr(34) & ") & vbCrLf" & vbCrLf
    CloseDB
    
    OpenDB "Select d_nome From detalhes Where d_exibir = 'S'"
    vContador = 0
    Do While Not TbDet.EOF
    If vContador = 6 Then: GoTo vai6d
    vContador = vContador + 1
vTexto = vTexto & "Response.Write (" & Chr(34) & "<td width=3% align=center height=25 colspan=3><font face=" & cmbFonte.Text & " size=" & cmbTamanho.Text & " color=#000000>" & Chr(34) & Chr(32) & Chr(38) & Chr(32) & "rsTemp(" & Chr(34) & TbDet("d_nome") & Chr(34) & ")" & Chr(32) & Chr(38) & Chr(32) & Chr(34) & "</font></td>" & Chr(34) & ")" & vbCrLf
    TbDet.MoveNext
    Loop
vai6d:
    CloseDB

vTexto = vTexto & "Response.Write (" & Chr(34) & "</tr>" & Chr(34) & ") & vbCrLf" & vbCrLf

vTexto = vTexto & "rsTemp.MoveNext" & vbCrLf
vTexto = vTexto & "If rsTemp.EOF Then Exit For" & vbCrLf
vTexto = vTexto & "Next" & vbCrLf
vTexto = vTexto & "End Sub" & vbCrLf

vTexto = vTexto & "Response.Write (" & Chr(34) & "</table>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</center>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</div>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<hr>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Navega" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<br>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "<p align=right><font size=2><a href=http://www.pagemaster.tk target=_blank>Página criada com PageMaster</a></font></p>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</body>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "Response.Write (" & Chr(34) & "</html>" & Chr(34) & ") & vbCrLf" & vbCrLf
vTexto = vTexto & "rsTemp.Close" & vbCrLf
vTexto = vTexto & "objDB.Close" & vbCrLf
vTexto = vTexto & "Set rsTemp = Nothing" & vbCrLf
vTexto = vTexto & "Set objDB = Nothing" & vbCrLf
vTexto = vTexto & "%>"

txtCriaConsulta.Text = vTexto
SaveFileAs Caminho & txtNomeArquivo.Text & ".asp", txtCriaConsulta
vTexto = ""
CriarInclusao
End Sub

Public Sub CriarInclusao()
length = GetPrivateProfileString( _
"Config", "Salvar", App.Path, _
buf, Len(buf), App.Path & "\pMaster.ini")
Caminho = Left$(buf, length) & "\"

vTexto = vTexto & "<%" & vbCrLf
vTexto = vTexto & "'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->" & vbCrLf
vTexto = vTexto & "'<------ /  Gerando com PageMaster   \ ------>" & vbCrLf
vTexto = vTexto & "'<----- /  http:\\www.pagemaster.tk   \ ----->" & vbCrLf
vTexto = vTexto & "'<----- \_____________________________/ ----->" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "'NOME DO ARQUIVO...: " & txtNomeArquivo.Text & "_INC.ASP" & vbCrLf
vTexto = vTexto & "'CRIADO EM.........: " & Now() & vbCrLf
vTexto = vTexto & "'---------------------------------------------" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "Dim vOpcao" & vbCrLf
vTexto = vTexto & "vOpcao = Request.Form(" & Chr(34) & "Opt" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "If vOpcao <> " & Chr(34) & "" & Chr(34) & " Then" & vbCrLf

vTexto = vTexto & "On Error Resume Next" & vbCrLf

    OpenDB "Select * From detalhes Where d_exibir = 'S'"
    Do While Not TbDet.EOF
    vCampo = Replace(TbDet("d_nome"), " ", "", 1)
vTexto = vTexto & "    If request.Form(" & Chr(34) & "f" & vCampo & Chr(34) & ") = " & Chr(34) & Chr(34) & " Then" & vbCrLf
vTexto = vTexto & "            mr" & vCampo & " = Null" & vbCrLf
vTexto = vTexto & "    Else" & vbCrLf
vTexto = vTexto & "            mr" & vCampo & " = request.Form(" & Chr(34) & "f" & vCampo & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "            mr" & vCampo & " = Replace(mr" & vCampo & ", Chr(13)," & Chr(34) & "<BR>" & Chr(34) & ", 1)" & vbCrLf
vTexto = vTexto & "    End If" & vbCrLf
vTexto = vTexto & "'--------------------------------------------------------------------------" & vbCrLf
    TbDet.MoveNext
    Loop
    CloseDB
    
'Abre a conexao com a base de dados
vTexto = vTexto & "%>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<!--#include file=" & Chr(34) & txtNomeArquivo.Text & "_CNX.asp" & Chr(34) & "-->" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<%" & vbCrLf

vTexto = vTexto & "Set rsTemp = Server.CreateObject(" & Chr(34) & "ADODB.Recordset" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "If request.Form(" & Chr(34) & "acao" & Chr(34) & ") = " & Chr(34) & "editar" & Chr(34) & " Then" & vbCrLf
vTexto = vTexto & "'=====================================================" & vbCrLf
vTexto = vTexto & "    sql = " & Chr(34) & "SELECT * FROM " & lblNomeTabela.Caption & Chr(34) & vbCrLf
    
    OpenDB "Select Top 1 d_nome From detalhes"
vTexto = vTexto & "    sql = sql & " & Chr(34) & " WHERE " & TbDet("d_nome") & " = " & Chr(34) & " & request.Form(" & Chr(34) & "fauto" & Chr(34) & ")" & vbCrLf
    CloseDB
    
vTexto = vTexto & "'=====================================================" & vbCrLf
vTexto = vTexto & "    rsTemp.Open sql, objDB, 3, 3" & vbCrLf
vTexto = vTexto & "Else" & vbCrLf
vTexto = vTexto & "    rsTemp.Open " & Chr(34) & lblNomeTabela.Caption & Chr(34) & ", objDB, 3, 3" & vbCrLf
vTexto = vTexto & "    rsTemp.AddNew" & vbCrLf
vTexto = vTexto & "End If" & vbCrLf
vTexto = vTexto & "" & vbCrLf

    OpenDB "Select * From detalhes Where d_exibir = 'S'"
    Do While Not TbDet.EOF
vTexto = vTexto & "rsTemp(" & Chr(34) & TbDet("d_nome") & Chr(34) & ")" & " = " & "mr" & TbDet("d_nome") & vbCrLf
    TbDet.MoveNext
    Loop
    CloseDB
    
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "rsTemp.Update" & vbCrLf
vTexto = vTexto & "rsTemp.Close" & vbCrLf
vTexto = vTexto & "objDB.Close" & vbCrLf

vTexto = vTexto & "If Err.Number > 0 Then" & vbCrLf
vTexto = vTexto & "    response.write " & Chr(34) & "Ocorreram erros de VBS, copie a tela e envie para <a href='mailto:pagemaster@com4.com.br'>Suporte de Sistemas</a>" & Chr(34) & " & vbCrLf" & vbCrLf
vTexto = vTexto & "    response.write " & Chr(34) & "Número = Err.Numberv " & Chr(34) & " & vbCrLf" & vbCrLf
vTexto = vTexto & "    response.write " & Chr(34) & "Descrição = Err.Description " & Chr(34) & " & vbCrLf" & vbCrLf
vTexto = vTexto & "    response.write " & Chr(34) & "Help Context = Err.HelpContext" & Chr(34) & " & vbCrLf" & vbCrLf
vTexto = vTexto & "    response.write " & Chr(34) & "Help Path = Err.helppath" & Chr(34) & " & vbCrLf" & vbCrLf
vTexto = vTexto & "    response.write " & Chr(34) & "Native Error = Err.nativeerror" & Chr(34) & " & vbCrLf" & vbCrLf
vTexto = vTexto & "    response.write " & Chr(34) & "Source = Err.Source" & Chr(34) & " & vbCrLf" & vbCrLf
vTexto = vTexto & "    response.write " & Chr(34) & "SQLState = Err.SQLState" & Chr(34) & " & vbCrLf" & vbCrLf
vTexto = vTexto & "End If" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "If objDB.Errors.Count > 0 Then" & vbCrLf
vTexto = vTexto & "    response.write " & Chr(34) & "Ocorreram erros no BD, copie a tela e envie para <a href='mailto:pagemaster@com4.com.br'>Suporte de Sistemas</a>" & Chr(34) & " & vbCrLf" & vbCrLf
vTexto = vTexto & "    For Counter = 0 To objDB.Errors.Count & vbCrLf" & vbCrLf
vTexto = vTexto & "        response.write " & Chr(34) & "Número    ->" & Chr(34) & " " & Chr(38) & " objDB.Errors(Counter).Number" & " " & Chr(38) & " " & Chr(34) & "<P>" & Chr(34) & vbCrLf
vTexto = vTexto & "        response.write " & Chr(34) & "Descrição ->" & Chr(34) & " " & Chr(38) & " objDB.Errors(Counter).Description" & " " & Chr(38) & " " & Chr(34) & "<P>" & Chr(34) & vbCrLf
vTexto = vTexto & "    Next" & vbCrLf
vTexto = vTexto & "Else" & vbCrLf
vTexto = vTexto & "        response.redirect " & Chr(34) & txtNomeArquivo & ".asp" & Chr(34) & vbCrLf

vTexto = vTexto & "End If" & vbCrLf

vTexto = vTexto & "Else" & vbCrLf
vTexto = vTexto & "%>" & vbCrLf
vTexto = vTexto & "" & vbCrLf

vTexto = vTexto & "<html>" & vbCrLf
vTexto = vTexto & "<head>" & vbCrLf
vTexto = vTexto & "<title>" & txtInclusao.Text & "</title>" & vbCrLf
vTexto = vTexto & "<meta name=generator content=http://www.pagemaster.tk>" & vbCrLf
vTexto = vTexto & "</head>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<body bgcolor=white text=black link=#000099 vlink=#000099 alink=#000099 leftmargin=0 marginwidth=0 topmargin=0 marginheight=0>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<script Language=" & Chr(34) & "JavaScript" & Chr(34) & ">" & vbCrLf
vTexto = vTexto & "function Validator(theForm)" & vbCrLf
vTexto = vTexto & "{" & vbCrLf

    OpenDB "Select * From detalhes Where d_nulo = 'N' And d_exibir = 'S'"
    Do While Not TbDet.EOF
    vCampo = Replace(TbDet("d_nome"), " ", "", 1)
vTexto = vTexto & "  if (theForm." & "f" & vCampo & ".value == " & Chr(34) & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "  {" & vbCrLf
vTexto = vTexto & "    alert(" & Chr(34) & "Digite um valor para o campo " & TbDet("d_rotulo") & Chr(34) & ");" & vbCrLf
vTexto = vTexto & "    theForm." & "f" & vCampo & ".focus();" & vbCrLf
vTexto = vTexto & "    return (false);" & vbCrLf
vTexto = vTexto & "  }" & vbCrLf
    TbDet.MoveNext
    Loop
    CloseDB
    
vTexto = vTexto & "  return (true);" & vbCrLf
vTexto = vTexto & "}" & vbCrLf
vTexto = vTexto & "</script>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "    <p align=center><font face=Arial size=4><b>" & txtInclusao.Text & "</b></font></p>" & vbCrLf
vTexto = vTexto & "<table align=center bgcolor=#ffe4b5>" & vbCrLf
vTexto = vTexto & "<form name=FrmAdm onsubmit=" & Chr(34) & "return Validator(this)" & Chr(34) & " method=post action=" & txtNomeArquivo & "_Inc.asp>" & vbCrLf
vTexto = vTexto & "  <TBODY>" & vbCrLf
vTexto = vTexto & "    <tr>" & vbCrLf
vTexto = vTexto & "        <td bgcolor=#D2B48C VALIGN=top colspan=2>" & vbCrLf
vTexto = vTexto & "                <p align=center>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<font color=#A52829 face=Verdana size=2><b>Incluir Novo Registro em " & lblNomeTabela.Caption & "</b></font></p></td>" & vbCrLf
vTexto = vTexto & "    </tr>" & vbCrLf

    Dim xRotulo As String
    OpenDB "Select * From detalhes Where d_exibir = 'S'"
    Do While Not TbDet.EOF
    xRotulo = TbDet("d_rotulo")
vTexto = vTexto & "<tr>" & vbCrLf
vTexto = vTexto & "        <td><p><font color=#8C4510 face=Verdana size=2>" & xRotulo & "</font></p></td>" & vbCrLf
    
    vCampo = Replace(TbDet("d_nome"), " ", "", 1)
    
    If TbDet("d_tipo") = "Campo de Texto" Then
    
    Dim vInicio As String
    If TbDet("d_valorInicial") = "" Then
        vInicio = ""
    Else
        vInicio = "value=" & TbDet("d_valorInicial")
    End If
        
        If TbDet("d_senha") = "N" Then
vTexto = vTexto & "        <td bgcolor=#fffaf0><p><input type=text name=f" & vCampo & " " & vInicio & " size=" & TbDet("d_tamaho") & " maxlength=" & TbDet("d_maximo") & "></p></td>" & vbCrLf
        Else
vTexto = vTexto & "        <td bgcolor=#fffaf0><p><input type=password name=f" & vCampo & " " & vInicio & " size=" & TbDet("d_tamaho") & " maxlength=" & TbDet("d_maximo") & "></p></td>" & vbCrLf
        End If
    Else
vTexto = vTexto & "        <td bgcolor=#fffaf0><p align=left><textarea name=f" & vCampo & " " & vInicio & " rows=" & TbDet("d_maximo") & " cols=" & TbDet("d_tamaho") & "></textarea></p></td>" & vbCrLf
    End If
vTexto = vTexto & "</tr>" & vbCrLf
    TbDet.MoveNext
    Loop
    CloseDB
    
vTexto = vTexto & "   </tr>" & vbCrLf
vTexto = vTexto & "    <tr>" & vbCrLf
vTexto = vTexto & "        <td colspan=2 height=16 valign=bottom>" & vbCrLf
vTexto = vTexto & "                <table align=center cellpadding=0 cellspacing=0 width=100% height=22 style=" & Chr(34) & "WIDTH: 100%" & Chr(34) & ">" & vbCrLf
vTexto = vTexto & "                    <tr>" & vbCrLf
vTexto = vTexto & "                        <td width=248 valign=center bgcolor=#D2B48C>" & vbCrLf
vTexto = vTexto & "                            <p align=center><input onclick=" & Chr(34) & "JavaScript:self.history.go(-1)" & Chr(34) & " type=button name=btVoltar value=" & Chr(34) & "  Voltar  " & Chr(34) & " style=" & Chr(34) & "FONT-SIZE: 8pt; FONT-FAMILY: Verdana" & Chr(34) & "><input type=hidden name=Opt value=new></p>" & vbCrLf
vTexto = vTexto & "                        </td>" & vbCrLf
vTexto = vTexto & "                        <td width=248 valign=center bgcolor=#D2B48C><p align=center><input type=submit name=cmdInclui value=" & Chr(34) & "Incluir Registro" & Chr(34) & " style=" & Chr(34) & "FONT-SIZE: 8pt; FONT-FAMILY: Verdana" & Chr(34) & "></p>" & vbCrLf
vTexto = vTexto & "                        </td>" & vbCrLf
vTexto = vTexto & "                    </tr>" & vbCrLf
vTexto = vTexto & "                </table>" & vbCrLf
vTexto = vTexto & "        </td>" & vbCrLf
vTexto = vTexto & "    </tr>" & vbCrLf
vTexto = vTexto & "</form></TBODY></table>" & vbCrLf
vTexto = vTexto & "</body>" & vbCrLf
vTexto = vTexto & "</html>" & vbCrLf
vTexto = vTexto & "<%" & vbCrLf
vTexto = vTexto & "End If" & vbCrLf
vTexto = vTexto & "%>"

txtCriaInclusao.Text = vTexto
SaveFileAs Caminho & txtNomeArquivo.Text & "_INC.asp", txtCriaInclusao
vTexto = ""
CriarAlteracao
End Sub

Public Sub CriarConexao()
If txtNomeArquivo.Text = "" Then
If Idioma = "BR" Then
MsgBox "Informe o nome do arquivo", vbExclamation, "Atenção!!"
Else
MsgBox "Inform the name of the file", vbExclamation, "Attention!!"
End If
ConfigASP
txtNomeArquivo.SetFocus
Exit Sub
End If

length = GetPrivateProfileString( _
"Config", "Salvar", App.Path, _
buf, Len(buf), App.Path & "\pMaster.ini")
Caminho = Left$(buf, length) & "\"

If Dir$(Caminho & txtNomeArquivo.Text & ".asp") <> "" Then
If Idioma = "BR" Then
    If MsgBox("O arquivo " & Caminho & txtNomeArquivo & ".asp" & Chr(10) & _
    "já existe. Deseja substitui-lo?", vbQuestion + vbYesNo) = vbNo Then: Exit Sub
Else
    If MsgBox("The file " & Caminho & txtNomeArquivo & ".asp" & Chr(10) & _
    "it already exists. Does want to substitute it?", vbQuestion + vbYesNo) = vbNo Then: Exit Sub
End If
End If

vTexto = vTexto & "<%" & vbCrLf
vTexto = vTexto & "'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->" & vbCrLf
vTexto = vTexto & "'<------ /  Gerando com PageMaster   \ ------>" & vbCrLf
vTexto = vTexto & "'<----- /  http:\\www.pagemaster.tk   \ ----->" & vbCrLf
vTexto = vTexto & "'<----- \_____________________________/ ----->" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "'NOME DO ARQUIVO...: " & txtNomeArquivo.Text & "_CNX.ASP" & vbCrLf
vTexto = vTexto & "'CRIADO EM.........: " & Now() & vbCrLf
vTexto = vTexto & "'---------------------------------------------" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "Set objDB = Server.CreateObject(" & Chr(34) & "ADODB.Connection" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "objDB.CursorLocation = 3" & vbCrLf
vTexto = vTexto & "'O caminho do BDD por ser substituido pela instrução ODBC " & vbCrLf
vTexto = vTexto & "sDBName =" & Chr(34) & "driver={Microsoft Access Driver (*.mdb)};dbq=" & txtBdd.Text & Chr(34) & vbCrLf
vTexto = vTexto & "objDB.Open sDBName" & vbCrLf
vTexto = vTexto & "%>" & vbCrLf

txtCriaConexao.Text = vTexto
SaveFileAs Caminho & txtNomeArquivo.Text & "_CNX.asp", txtCriaConexao
vTexto = ""
CriarConsulta

End Sub

Public Sub CriarAlteracao()
length = GetPrivateProfileString( _
"Config", "Salvar", App.Path, _
buf, Len(buf), App.Path & "\pMaster.ini")
Caminho = Left$(buf, length) & "\"

vTexto = vTexto & "<%" & vbCrLf
vTexto = vTexto & "'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->" & vbCrLf
vTexto = vTexto & "'<------ /  Gerando com PageMaster   \ ------>" & vbCrLf
vTexto = vTexto & "'<----- /  http:\\www.pagemaster.tk   \ ----->" & vbCrLf
vTexto = vTexto & "'<----- \_____________________________/ ----->" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "'NOME DO ARQUIVO...: " & txtNomeArquivo.Text & "_ALT.ASP" & vbCrLf
vTexto = vTexto & "'CRIADO EM.........: " & Now() & vbCrLf
vTexto = vTexto & "'---------------------------------------------" & vbCrLf
vTexto = vTexto & "%>" & vbCrLf

'Abre a conexao com a base de dados

vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<!--#include file=" & Chr(34) & txtNomeArquivo.Text & "_CNX.asp" & Chr(34) & "-->" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<%" & vbCrLf
vTexto = vTexto & "On Error Resume Next" & vbCrLf
vTexto = vTexto & "Dim vCod" & vbCrLf
vTexto = vTexto & "vCod = Request.QueryString(" & Chr(34) & "id" & Chr(34) & ")" & vbCrLf

    OpenDB "Select Top 1 d_nome From detalhes"
vTexto = vTexto & "    sql = " & Chr(34) & "SELECT * FROM " & lblNomeTabela.Caption & " Where " & TbDet("d_nome") & " = " & Chr(34) & " & vCod" & vbCrLf
    CloseDB
    
vTexto = vTexto & "Set rsTemp = Server.CreateObject(" & Chr(34) & "ADODB.Recordset" & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "    rsTemp.Open sql, objDB, 3, 3" & vbCrLf
vTexto = vTexto & "%>" & vbCrLf

vTexto = vTexto & "<html>" & vbCrLf
vTexto = vTexto & "<head>" & vbCrLf
vTexto = vTexto & "<title>" & txtAlteracao.Text & "</title>" & vbCrLf
vTexto = vTexto & "<meta name=generator content=http://www.pagemaster.tk>" & vbCrLf
vTexto = vTexto & "</head>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<body bgcolor=white text=black link=#000099 vlink=#000099 alink=#000099 leftmargin=0 marginwidth=0 topmargin=0 marginheight=0>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<script Language=" & Chr(34) & "JavaScript" & Chr(34) & ">" & vbCrLf
vTexto = vTexto & "function Validator(theForm)" & vbCrLf
vTexto = vTexto & "{" & vbCrLf

    OpenDB "Select * From detalhes Where d_nulo = 'N' And d_exibir = 'S'"
    Do While Not TbDet.EOF
    vCampo = Replace(TbDet("d_nome"), " ", "", 1)
vTexto = vTexto & "  if (theForm." & "f" & vCampo & ".value == " & Chr(34) & Chr(34) & ")" & vbCrLf
vTexto = vTexto & "  {" & vbCrLf
vTexto = vTexto & "    alert(" & Chr(34) & "Digite um valor para o campo " & TbDet("d_rotulo") & Chr(34) & ");" & vbCrLf
vTexto = vTexto & "    theForm." & "f" & vCampo & ".focus();" & vbCrLf
vTexto = vTexto & "    return (false);" & vbCrLf
vTexto = vTexto & "  }" & vbCrLf
    TbDet.MoveNext
    Loop
    CloseDB
    
vTexto = vTexto & "  return (true);" & vbCrLf
vTexto = vTexto & "}" & vbCrLf
vTexto = vTexto & "</script>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "    <p align=center><font face=Arial size=4><b>" & txtAlteracao.Text & "</b></font></p>" & vbCrLf
vTexto = vTexto & "<table align=center bgcolor=#ffe4b5>" & vbCrLf
vTexto = vTexto & "<form name=FrmAdm onsubmit=" & Chr(34) & "return Validator(this)" & Chr(34) & " method=post action=" & txtNomeArquivo & "_Inc.asp>" & vbCrLf

vTexto = vTexto & "<input type=hidden name=acao value=editar>" & vbCrLf
vTexto = vTexto & "<input type=hidden name=fauto value=<%=vCod%>>" & vbCrLf

vTexto = vTexto & "  <TBODY>" & vbCrLf
vTexto = vTexto & "    <tr>" & vbCrLf
vTexto = vTexto & "        <td bgcolor=#D2B48C VALIGN=top colspan=2>" & vbCrLf
vTexto = vTexto & "                <p align=center>" & vbCrLf
vTexto = vTexto & "" & vbCrLf
vTexto = vTexto & "<font color=#A52829 face=Verdana size=2><b>Alterando Registro em " & lblNomeTabela.Caption & "</b></font></p></td>" & vbCrLf
vTexto = vTexto & "    </tr>" & vbCrLf


    Dim xRotulo As String, xMr1 As String, xMr2 As String
    OpenDB "Select * From detalhes Where d_exibir = 'S'"
    Do While Not TbDet.EOF
    xRotulo = TbDet("d_rotulo")
vTexto = vTexto & "<tr>" & vbCrLf
vTexto = vTexto & "        <td><p><font color=#8C4510 face=Verdana size=2>" & xRotulo & "</font></p></td>" & vbCrLf
    
    vCampo = Replace(TbDet("d_nome"), " ", "", 1)
    xMr1 = "mr" & vCampo & " = Replace(rsTemp(" & Chr(34) & vCampo & Chr(34) & "), " & Chr(34) & "<BR>" & Chr(34) & ", chr(13), 1)"
    xMr2 = "mr" & vCampo & " = Replace(" & "mr" & vCampo & ", " & Chr(34) & "''" & Chr(34) & ", " & Chr(34) & "'" & Chr(34) & ", 1)"
    
vTexto = vTexto & "<%" & vbCrLf
vTexto = vTexto & xMr1 & vbCrLf
vTexto = vTexto & xMr2 & vbCrLf
vTexto = vTexto & "%>"
    Dim vInicio2 As String
    If TbDet("d_tipo") = "Campo de Texto" Then
    
        vInicio = " value='<%=mr" & vCampo & "%>'"
        vInicio2 = "'<%=mr" & vCampo & "%>'"
 
        If TbDet("d_senha") = "N" Then
vTexto = vTexto & "        <td bgcolor=#fffaf0><p><input type=text name=f" & vCampo & " size=" & TbDet("d_tamaho") & " maxlength=" & TbDet("d_maximo") & " " & vInicio & "></p></td>" & vbCrLf
        Else
vTexto = vTexto & "        <td bgcolor=#fffaf0><p><input type=password name=f" & vCampo & " size=" & TbDet("d_tamaho") & " maxlength=" & TbDet("d_maximo") & " " & vInicio & "></p></td>" & vbCrLf
        End If
    
    Else
        vInicio = "<%=mr" & vCampo & "%>"
vTexto = vTexto & "        <td bgcolor=#fffaf0><p align=left><textarea name=f" & vCampo & " rows=" & TbDet("d_maximo") & " cols=" & TbDet("d_tamaho") & ">" & vInicio2 & "</textarea></p></td>" & vbCrLf
    End If
    
vTexto = vTexto & "</tr>" & vbCrLf
    TbDet.MoveNext
    Loop
    CloseDB
    
vTexto = vTexto & "   </tr>" & vbCrLf
vTexto = vTexto & "    <tr>" & vbCrLf
vTexto = vTexto & "        <td colspan=2 height=16 valign=bottom>" & vbCrLf
vTexto = vTexto & "                <table align=center cellpadding=0 cellspacing=0 width=100% height=22 style=" & Chr(34) & "WIDTH: 100%" & Chr(34) & ">" & vbCrLf
vTexto = vTexto & "                    <tr>" & vbCrLf
vTexto = vTexto & "                        <td width=248 valign=center bgcolor=#D2B48C>" & vbCrLf
vTexto = vTexto & "                            <p align=center><input onclick=" & Chr(34) & "JavaScript:self.history.go(-1)" & Chr(34) & " type=button name=btVoltar value=" & Chr(34) & "  Voltar  " & Chr(34) & " style=" & Chr(34) & "FONT-SIZE: 8pt; FONT-FAMILY: Verdana" & Chr(34) & "><input type=hidden name=Opt value=new></p>" & vbCrLf
vTexto = vTexto & "                        </td>" & vbCrLf
vTexto = vTexto & "                        <td width=248 valign=center bgcolor=#D2B48C><p align=center><input type=submit name=cmdInclui value=" & Chr(34) & "Alterar Registro" & Chr(34) & " style=" & Chr(34) & "FONT-SIZE: 8pt; FONT-FAMILY: Verdana" & Chr(34) & "></p>" & vbCrLf
vTexto = vTexto & "                        </td>" & vbCrLf
vTexto = vTexto & "                    </tr>" & vbCrLf
vTexto = vTexto & "                </table>" & vbCrLf
vTexto = vTexto & "        </td>" & vbCrLf
vTexto = vTexto & "    </tr>" & vbCrLf
vTexto = vTexto & "</form></TBODY></table>" & vbCrLf
vTexto = vTexto & "</body>" & vbCrLf
vTexto = vTexto & "</html>" & vbCrLf

vTexto = vTexto & "<%" & vbCrLf
vTexto = vTexto & "rsTemp.Close" & vbCrLf
vTexto = vTexto & "objDB.Close" & vbCrLf
vTexto = vTexto & "Set rsTemp = Nothing" & vbCrLf
vTexto = vTexto & "Set objDB = Nothing" & vbCrLf
vTexto = vTexto & "%>" & vbCrLf

freForms.Top = 600
freForms.Left = 105
freForms.Visible = True

freConfig.Visible = False
freCampos.Visible = False
freConfCampo.Visible = False

txtCriaAlteracao.Text = vTexto
SaveFileAs Caminho & txtNomeArquivo.Text & "_ALT.asp", txtCriaAlteracao
vTexto = ""
If Idioma = "BR" Then
MsgBox "O arquivo foi criado em " & Chr(10) & _
Caminho, vbInformation
Else
MsgBox "The file was created in " & Chr(10) & _
Caminho, vbInformation
End If

End Sub

Public Sub CarregaStrings()
    Dim Idioma As String
    length = GetPrivateProfileString( _
        "Config", "Idioma", App.Path, _
        buf, Len(buf), App.Path & "\pMaster.ini")
        Idioma = Left$(buf, length)

If Idioma = "BR" Then
lblDefineMail.Caption = LoadResString(118)
lblConfigMail.Caption = LoadResString(119)
lblGeraMail.Caption = LoadResString(120)
lblInfo.Caption = LoadResString(101)
freConfig.Caption = LoadResString(103)
Label37.Caption = LoadResString(104)
Label2.Caption = LoadResString(105)
Label6.Caption = LoadResString(106)
Label5.Caption = LoadResString(107)
Label3.Caption = LoadResString(108)
Label4.Caption = LoadResString(109)
Label7.Caption = LoadResString(110)
Label36.Caption = LoadResString(111)
Check1.Caption = LoadResString(112)
Check2.Caption = LoadResString(113)
Label13(0).Caption = LoadResString(114)
Label13(1).Caption = LoadResString(115)
lblNovo.Caption = LoadResString(116)
lblAbrir.Caption = LoadResString(117)
lblDefinir.Caption = LoadResString(118)
lblConfig.Caption = LoadResString(119)
lblGerar.Caption = LoadResString(120)
lblSair.Caption = LoadResString(121)
freOpt.Caption = LoadResString(122)
optAsp.Caption = LoadResString(123)
optMail.Caption = LoadResString(124)
optCdonts.Caption = LoadResString(125)
freCampos.Caption = LoadResString(126)
freConfigAspMail.Caption = LoadResString(127)
Label27.Caption = LoadResString(128)
Label29.Caption = LoadResString(129)
Label30.Caption = LoadResString(130)
Label31.Caption = LoadResString(131)
Label32.Caption = LoadResString(132)
Label35.Caption = LoadResString(133)
Label33.Caption = LoadResString(134)
Label34.Caption = LoadResString(135)
freConfCampo.Caption = LoadResString(136)
Label8.Caption = LoadResString(137)
ckNulo.Caption = LoadResString(138)
ckExibir.Caption = LoadResString(139)
Label9.Caption = LoadResString(140)
Label11.Caption = LoadResString(141)
Label10.Caption = LoadResString(142)
Label12.Caption = LoadResString(143)
optSenha(0).Caption = LoadResString(144)
optSenha(1).Caption = LoadResString(145)
Label14.Caption = LoadResString(146)
Label15.Caption = LoadResString(147)
freAspMail.Caption = LoadResString(148)
Label16.Caption = LoadResString(150)
Label17.Caption = LoadResString(151)
ckObrigatorio.Caption = LoadResString(152)
ckSenha.Caption = LoadResString(153)
Label21.Caption = LoadResString(154)
Label22.Caption = LoadResString(155)
Label23.Caption = LoadResString(156)
Check3.Caption = LoadResString(157)
lblMail(0).Caption = LoadResString(158)
lblMail(1).Caption = LoadResString(159)
lblMail(2).Caption = LoadResString(160)
lblMail(3).Caption = LoadResString(161)
freForms.Caption = LoadResString(162)
lblTConexao.Caption = LoadResString(163)
lblTConsulta.Caption = LoadResString(164)
lblTInclusao.Caption = LoadResString(165)
lblTAlteracao.Caption = LoadResString(166)
freGeradoAsp.Caption = LoadResString(167)
lblInserirDados.Caption = LoadResString(169)
lblEnviarDados.Caption = LoadResString(170)
mnuArquivo(0).Caption = LoadResString(171)
mnuConfig.Caption = LoadResString(172)
mnuAjuda(0).Caption = LoadResString(173)
mnuarq(0).Caption = LoadResString(116)
mnuarq(2).Caption = LoadResString(121)
mnuConf(0).Caption = LoadResString(174)
mnuConf(1).Caption = LoadResString(175)
mnucont(0).Caption = LoadResString(176)
mnucont(1).Caption = LoadResString(177)
mnucont(4).Caption = LoadResString(295)
Label19.Caption = LoadResString(279)
Label20.Caption = LoadResString(281)
Frame1.Caption = LoadResString(283)
Label25.Caption = LoadResString(285)
freRegistro.Caption = LoadResString(288)
Label1.Caption = LoadResString(290)
Label38.Caption = LoadResString(291)
Label39.Caption = LoadResString(293)
lblRegistro.Caption = "Registrar"
Else
lblRegistro.Caption = LoadResString(287)
Label39.Caption = LoadResString(294)
Label38.Caption = LoadResString(292)
Label1.Caption = LoadResString(289)
freRegistro.Caption = LoadResString(287)
Label25.Caption = LoadResString(286)
Frame1.Caption = LoadResString(284)
Label20.Caption = LoadResString(282)
Label19.Caption = LoadResString(280)
mnucont(1).Caption = LoadResString(187)
mnucont(0).Caption = LoadResString(188)
mnuConf(1).Caption = LoadResString(189)
mnuConf(0).Caption = LoadResString(190)
mnucont(4).Caption = LoadResString(296)
mnuarq(2).Caption = LoadResString(191)
mnuarq(0).Caption = LoadResString(192)
mnuAjuda(0).Caption = LoadResString(193)
mnuConfig.Caption = LoadResString(194)
mnuArquivo(0).Caption = LoadResString(195)
lblEnviarDados.Caption = LoadResString(196)
lblInserirDados.Caption = LoadResString(197)
freGeradoAsp.Caption = LoadResString(198)
lblTAlteracao.Caption = LoadResString(199)
lblTInclusao.Caption = LoadResString(200)
lblTConsulta.Caption = LoadResString(201)
lblTConexao.Caption = LoadResString(202)
freForms.Caption = LoadResString(203)
lblMail(3).Caption = LoadResString(204)
lblMail(2).Caption = LoadResString(205)
lblMail(1).Caption = LoadResString(206)
lblMail(0).Caption = LoadResString(207)
Check3.Caption = LoadResString(208)
Label23.Caption = LoadResString(209)
Label22.Caption = LoadResString(210)
Label21.Caption = LoadResString(211)
ckSenha.Caption = LoadResString(212)
ckObrigatorio.Caption = LoadResString(213)
Label17.Caption = LoadResString(214)
Label16.Caption = LoadResString(215)
freAspMail.Caption = LoadResString(216)
Label15.Caption = LoadResString(218)
Label14.Caption = LoadResString(219)
optSenha(1).Caption = LoadResString(220)
optSenha(0).Caption = LoadResString(221)
Label12.Caption = LoadResString(222)
Label10.Caption = LoadResString(223)
Label11.Caption = LoadResString(224)
Label9.Caption = LoadResString(225)
ckExibir.Caption = LoadResString(226)
ckNulo.Caption = LoadResString(227)
Label8.Caption = LoadResString(228)
freConfCampo.Caption = LoadResString(229)
Label34.Caption = LoadResString(230)
Label33.Caption = LoadResString(231)
Label35.Caption = LoadResString(232)
Label32.Caption = LoadResString(233)
Label31.Caption = LoadResString(234)
Label30.Caption = LoadResString(235)
Label29.Caption = LoadResString(236)
Label27.Caption = LoadResString(237)
freConfigAspMail.Caption = LoadResString(238)
freCampos.Caption = LoadResString(239)
optCdonts.Caption = LoadResString(240)
optMail.Caption = LoadResString(241)
freOpt.Caption = LoadResString(242)
optAsp.Caption = LoadResString(243)
lblSair.Caption = LoadResString(244)
lblGerar.Caption = LoadResString(245)
lblConfig.Caption = LoadResString(246)
lblDefinir.Caption = LoadResString(247)
lblAbrir.Caption = LoadResString(248)
lblNovo.Caption = LoadResString(192)
Label13(1).Caption = LoadResString(249)
Label13(0).Caption = LoadResString(250)
Check2.Caption = LoadResString(251)
Check1.Caption = LoadResString(252)
Label36.Caption = LoadResString(253)
Label7.Caption = LoadResString(254)
Label4.Caption = LoadResString(255)
Label3.Caption = LoadResString(256)
Label5.Caption = LoadResString(257)
Label6.Caption = LoadResString(258)
Label2.Caption = LoadResString(259)
Label37.Caption = LoadResString(260)
freConfig.Caption = LoadResString(261)
lblInfo.Caption = LoadResString(262)
lblDefineMail.Caption = LoadResString(247)
lblConfigMail.Caption = LoadResString(246)
lblGeraMail.Caption = LoadResString(245)
End If

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

Public Sub registrar()
If txtChaveRegistro.Text = "" Then: Exit Sub
  ActiveLock.LiberationKey = txtChaveRegistro.Text
  If ActiveLock.RegisteredUser Then
    freRegistro.Visible = False
    MsgBox "Obrigado por se registrar!", vbExclamation
  Else
    MsgBox "O número da chave não é válido!", vbCritical
    txtChaveRegistro.SelStart = 0
    txtChaveRegistro.SelLength = Len(txtChaveRegistro)
    txtChaveRegistro.SetFocus
  End If
End Sub

Public Sub obter()
frmRegistro.Show vbModal
End Sub

Private Sub RemoveMenus(Frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(hwnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

