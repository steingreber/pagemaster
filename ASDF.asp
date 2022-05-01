<html>
<head>
<title>ASDF</title>
</head>
<body>
<center><font face=Arial size=4><b>Formulário AspMail</b></font></h2></center>
<script Language="JavaScript">
function Validator(theForm)
{
  if (theForm.tNnomedocliente.value == "")
  {
    alert("Digite um valor para o campo NOME DO CLIENTE:");
    theForm.tNnomedocliente.focus();
    return (false);
  }
  if (theForm.tEendereco.value == "")
  {
    alert("Digite um valor para o campo ENDERECO:");
    theForm.tEendereco.focus();
    return (false);
  }
  if (theForm.tEestado.value == "")
  {
    alert("Digite um valor para o campo ESTADO:");
    theForm.tEestado.focus();
    return (false);
  }
  if (theForm.tCcep.value == "")
  {
    alert("Digite um valor para o campo CEP:");
    theForm.tCcep.focus();
    return (false);
  }
  if (theForm.tTtelefoneparacon.value == "")
  {
    alert("Digite um valor para o campo TELEFONE PARA CONTATO:");
    theForm.tTtelefoneparacon.focus();
    return (false);
  }
  if (theForm.tCcnpj.value == "")
  {
    alert("Digite um valor para o campo C N P J:");
    theForm.tCcnpj.focus();
    return (false);
  }
  if (theForm.tIinscestadual.value == "")
  {
    alert("Digite um valor para o campo INSC ESTADUAL:");
    theForm.tIinscestadual.focus();
    return (false);
  }
  return (true);
}
</script>
        <form method=POST name=Manda onsubmit="return Validator(this)" Action="ASDFE.asp">
          <table border=0 width=100% cellspacing=0 cellpadding=0>
            <tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>NOME DO CLIENTE:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tNnomedocliente size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>ENDERECO:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tEendereco size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>CIDADE:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tCcidade size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>ESTADO:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tEestado size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>CEP:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tCcep size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>TELEFONE PARA CONTATO:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tTtelefoneparacon size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>C N P J:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tCcnpj size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>INSC ESTADUAL:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tIinscestadual size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>EMAIL:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tEemail size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>PAGINA WEB:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tPpaginaweb size=52></td>
              </tr>
             <tr>
         <td width=760 colspan=2>
          <p align=center><input type=submit value=Enviar name=cmdEnvio></p>
         </td>
             </tr>
            </center>
              <center>
            </table>
          </form>
</body>
