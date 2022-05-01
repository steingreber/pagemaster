Attribute VB_Name = "Global"
Option Explicit
'-------Cancela o X do form
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Const MF_BYPOSITION = &H400&
Public ReadyToClose As Boolean
'---------------------------
Private Declare Function GetVolumeInformation Lib "kernel32" _
     Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
     ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
     lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
     lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
     ByVal nFileSystemNameSize As Long) As Long

Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public NomeArq As String, PathApp As String
Public Banco As DAO.Database
Public TbDet As Recordset, Tb As Recordset
Public vEdita As String, Idioma As String, NomeTabela As String

Public Sub abrirBanco()
On Error GoTo errbanco
    NomeArq = "\campos.mdb"
    PathApp = App.Path & NomeArq
    NomeTabela = "detalhes"
    Set Banco = OpenDatabase(PathApp)
    Set TbDet = Banco.OpenRecordset(NomeTabela, dbOpenDynaset)
Exit Sub
errbanco:
MsgBox Err & " " & Err.Description, vbInformation, "Erro!"
End Sub

'RPad - Enche caracteres à direita de uma string
Public Function RPad(St As String, Tm As Integer, Ch As String) As String
 Dim X As String                                            'dimensiona
 If VarType(St) = vbString Then                             'se veio uma string
  X$ = St                                         'pega ela...
 Else                                             'senão,
  X$ = Str$(St)                                   'transforma em string
 End If
 RPad$ = Left$(LTrim$(X$) + String$(Tm, Ch$), Tm) 'completa com brancos à direita
End Function

'LPad - Enche caracteres à esquerda de uma string
Public Function LPad(St As Variant, Tm As Integer, Ch As String) As String
 Dim X As String                                            'dimensiona
 If VarType(St) = vbString Then                             'se veio uma string
  X$ = St                                         'pega ela...
 Else                                             'senão,
  X$ = Str$(St)                                   'transforma em string
 End If
 LPad$ = Right$(String$(Tm, Ch$) + LTrim$(X$), Tm) 'completa com brancos à esquerda
End Function

Public Sub OpenDB(tabela As String)

    NomeArq = "\campos.mdb"
    PathApp = App.Path & NomeArq
    NomeTabela = tabela
    Set Banco = OpenDatabase(PathApp)
    Set TbDet = Banco.OpenRecordset(NomeTabela, dbOpenDynaset)

End Sub

Public Sub CloseDB()

    TbDet.Close
    Banco.Close
    Set TbDet = Nothing
    Set Banco = Nothing

End Sub

Public Function DriveSerialNumber(strDrive As String) As String

Dim X As Long, lngSerialNum As Long

Dim strRoot As String
strRoot = Left$(strDrive, 1) & ":\"

X = GetVolumeInformation(strRoot, "", 12, lngSerialNum, 13, 14, "", 15)

DriveSerialNumber = Hex$(lngSerialNum)

End Function

