<%
html = ""
html = html & "<html>"
html = html & "<head>"
html = html & "<title>ASDF</title>"
html = html & "</head>"
Response.Write html

'<------- /°°°°°°°°°°°°°°°°°°°°°°°°°\ ------->
'<------ /  Gerando com PageMaster   \ ------>
'<----- /  http:\\www.pagemaster.tk   \ ----->
'<----- \_____________________________/ ----->

    Dim sMsgErr

    '***** Executa as ações desta página *****
    Sub ProcessaPagina()
    '***** Declaração das Variáveis *****
    Dim vNnomedocliente
    Dim vEendereco
    Dim vCcidade
    Dim vEestado
    Dim vCcep
    Dim vTtelefoneparacon
    Dim vCcnpj
    Dim vIinscestadual
    Dim vEemail
    Dim vPpaginaweb
    sEMail = "sadf" 'quando quero que a mensagem venha para mim

    '***** Obtém valores preenchidos no Formulário *****
    vNnomedocliente = Request.Form("tNnomedocliente")
    vEendereco = Request.Form("tEendereco")
    vCcidade = Request.Form("tCcidade")
    vEestado = Request.Form("tEestado")
    vCcep = Request.Form("tCcep")
    vTtelefoneparacon = Request.Form("tTtelefoneparacon")
    vCcnpj = Request.Form("tCcnpj")
    vIinscestadual = Request.Form("tIinscestadual")
    vEemail = Request.Form("tEemail")
    vPpaginaweb = Request.Form("tPpaginaweb")

    '***** Monta corpo da mensagem a enviar por e-mail
    sBodyText = sBodyText & "NOME DO CLIENTE: " & vNnomedocliente & vbCrLf   'corpo
    sBodyText = sBodyText & "ENDERECO: " & vEendereco & vbCrLf   'corpo
    sBodyText = sBodyText & "CIDADE: " & vCcidade & vbCrLf   'corpo
    sBodyText = sBodyText & "ESTADO: " & vEestado & vbCrLf   'corpo
    sBodyText = sBodyText & "CEP: " & vCcep & vbCrLf   'corpo
    sBodyText = sBodyText & "TELEFONE PARA CONTATO: " & vTtelefoneparacon & vbCrLf   'corpo
    sBodyText = sBodyText & "C N P J: " & vCcnpj & vbCrLf   'corpo
    sBodyText = sBodyText & "INSC ESTADUAL: " & vIinscestadual & vbCrLf   'corpo
    sBodyText = sBodyText & "EMAIL: " & vEemail & vbCrLf   'corpo
    sBodyText = sBodyText & "PAGINA WEB: " & vPpaginaweb & vbCrLf   'corpo
    sBodyText = sBodyText & "----------------------------------------------" & vbCrLf 'corpo

    '***** Envia E-Mail para o destinatário *****
    Set Mail = Server.CreateObject("Persits.MailSender")
    Mail.Host = "asdf" ' Especifique o nome do seu servidor SMTP.
    Mail.From = "asdf" ' Remetente da mensagem
    Mail.FromName = "SDAF" ' Nome do remetente
    Mail.AddAddress sEMail ' Destinatario da mensagem para
    Mail.Subject = "ASDF" 'assunto
    Mail.Body = sBodyText 'corpo da mensagem montada acima
    sMsgErr = ""
    On Error Resume Next
    Mail.Send
    If Err <> 0 Then
    sMsgErr = "Ocorreu o seguinte erro ao tentar enviar o e-mail: " & Err.Description
    End If
    On Error GoTo 0
    End Sub

    '***** Executa as ações desta página *****
    ProcessaPagina
html = ""
If (sMsgErr <> "") Then
html = html & "<body>"
html = html & "<div>"
html = html & "  <table width=100% border=0 cellspacing=0 cellpadding=0 height=21>"
html = html & "    <tr>"
html = html & "      <td height=23>"
html = html & "        <p><font size=1 face=Arial color=#000000> & =sMsgErr & <br>"
Else
html = html & "</font>"
html = html & "    </tr>"
html = html & "    <tr>"
html = html & "      <td height=101>"

html = html & "    </tr>"
html = html & "    <tr>"
html = html & "      <td height=101>"
html = html & "        <p align=center><font face=Verdana size=2 color=#008080>"
html = html & "          Sua mensagem foi encaminhada com sucesso.<br>"
html = html & "          Logo entraremos em contato. </font> <font face=Verdana size=2 color=#800000><br>"
html = html & "          Agradecemos por seu interesse!<br>"
html = html & "          <a href="javascript:close()">Fechar</a><br>"
html = html & "        </font> </p>"
End If
html = html & "    </tr>"
html = html & "  </table>"
html = html & "</div>"
html = html & "</body>"
html = html & "</html>"
Response.Write html
%>
