<html>
<head>
<title>PAPAGAIO</title>
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
  if (theForm.tCcep.value == "")
  {
    alert("Digite um valor para o campo CEP:");
    theForm.tCcep.focus();
    return (false);
  }
  if (theForm.tEestado.value == "")
  {
    alert("Digite um valor para o campo ESTADO:");
    theForm.tEestado.focus();
    return (false);
  }
  if (theForm.tEemail.value == "")
  {
    alert("Digite um valor para o campo EMAIL:");
    theForm.tEemail.focus();
    return (false);
  }
  return (true);
}
</script>
        <form method=POST name=Manda onsubmit="return Validator(this)" Action="ERTASE.asp">
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
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>CEP:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tCcep size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>ESTADO:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tEestado size=52></td>
              </tr>
              <td width="28%"><p align=right style="margin-right: 5"><font color=#000000 face=Verdana size=2><b>EMAIL:</b></font></td>
              <center>
              <td width="72%"><input type=text name=tEemail size=52></td>
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

