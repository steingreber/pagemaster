VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmCaminho 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caminho"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CDOpt 
      Left            =   315
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picAbrir 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4065
      MouseIcon       =   "frmCaminho.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmCaminho.frx":030A
      ScaleHeight     =   510
      ScaleWidth      =   1020
      TabIndex        =   3
      Top             =   1440
      Width           =   1020
      Begin VB.Label lblAbrir 
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
         Left            =   615
         MouseIcon       =   "frmCaminho.frx":0C09
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   150
         Width           =   255
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   165
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmCaminho.frx":0F13
      Top             =   630
      Width           =   4725
   End
   Begin VB.CommandButton cmdProc 
      Caption         =   "..."
      Height          =   585
      Index           =   0
      Left            =   4920
      TabIndex        =   1
      Top             =   615
      Width           =   255
   End
   Begin VB.Label lblaminho 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   180
      TabIndex        =   0
      Top             =   390
      Width           =   555
   End
End
Attribute VB_Name = "frmCaminho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Copia As String
Dim buf As String * 256
Dim length As Long

Private Type tProcuraInformação
    hWndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Function SHBrowseForFolder Lib "Shell32.dll" (bBrowse As tProcuraInformação) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" (ByVal lL_Item As Long, ByVal sDir As String) As Long

Private Sub cmdProc_Click(Index As Integer)
sDiretorio = sProcuraPorDiretório("Diretório onde serão armazenadas suas páginas ASP")
Text1.Text = sDiretorio
End Sub

Private Sub Form_Load()

If Idioma = "BR" Then
    lblaminho.Caption = LoadResString(102)
    Me.Caption = "Caminho"
Else
    lblaminho.Caption = LoadResString(276)
    Me.Caption = LoadResString(190)
End If

    length = GetPrivateProfileString( _
        "Config", "Salvar", App.Path, _
        buf, Len(buf), App.Path & "\pMaster.ini")
        Text1.Text = Left$(buf, length)

End Sub

Private Sub lblAbrir_Click()
gravaCaminho
End Sub

Private Sub picAbrir_Click()
gravaCaminho
End Sub

Public Function sProcuraPorDiretório(sTitulo As String) As String
On Error Resume Next
'ACIONA O BROWSER A PROCURA DE DIRETÓRIO

Dim oProcuraInformação      As tProcuraInformação
Dim lItem                   As Long
Dim sNomeDiretório          As String
   
oProcuraInformação.hWndOwner = hWnd
oProcuraInformação.pidlRoot = 0
oProcuraInformação.sDisplayName = Space$(260)
oProcuraInformação.sTitle = sTitulo
oProcuraInformação.ulFlags = 1 ' Retorna nome do diretorio.
oProcuraInformação.lpfn = 0
oProcuraInformação.lParam = 0
oProcuraInformação.iImage = 0

lItem = SHBrowseForFolder(oProcuraInformação)
If lItem Then
    sNomeDiretório = Space$(260)
    If SHGetPathFromIDList(lItem, sNomeDiretório) Then
        sProcuraPorDiretório = Left(sNomeDiretório, InStr(sNomeDiretório, Chr$(0)) - 1)
    Else
        sProcuraPorDiretório = ""
    End If
End If
End Function

Public Sub CopiaImages()
Dim d As String, e As String
d$ = App.Path & "\images\delete.gif"
e$ = App.Path & "\images\edit.gif"
FileCopy d$, Copia$ & "\delete.gif"
FileCopy e$, Copia$ & "\edit.gif"
End Sub

Public Sub gravaCaminho()
On Error GoTo ErrAbrir
    WritePrivateProfileString _
        "Config", "Salvar", _
        Text1, App.Path & "\pMaster.ini"
        
    length = GetPrivateProfileString( _
        "Config", "Salvar", App.Path, _
        buf, Len(buf), App.Path & "\pMaster.ini")
        
    Copia = Left$(buf, length) & "\images"
    
    If Idioma = "BR" Then
        If MsgBox("Será copiado para " & Chr(10) & Copia & " as seguintes imagens:" & Chr(10) _
        & Chr(10) & "delete.gif  e  edit.gif " & Chr(10) _
        & Chr(10) & "Esta pasta deverá ser copiada juntamente com seu arquivos!" _
        & Chr(10) & "Confirma esta ação?", vbQuestion + vbYesNo, "Atenção!!") = vbNo Then: Exit Sub
    Else
        If MsgBox("It will be copied for " & Chr(10) & Copia & " the following images:" & Chr(10) _
        & Chr(10) & "delete.gif  e  edit.gif " & Chr(10) _
        & Chr(10) & "This paste should be copied together with its files!" _
        & Chr(10) & "Does it confirm this action?", vbQuestion + vbYesNo, "Attention!!") = vbNo Then: Exit Sub
    End If
    
    MkDir Copia
    CopiaImages

    Unload Me
    Exit Sub
ErrAbrir:
If Err = 75 Then: CopiaImages: Unload Me

End Sub
