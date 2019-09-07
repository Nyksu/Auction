<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var ErrorMsg=""

if (String(Session("backurl"))=="undefined"){Session("backurl")="start.html"}

isFirst=String(Request.Form("login"))=="undefined"
if(!isFirst){	
	Pass=TextFormData(Request.Form("pass"),"")
	Nik=TextFormData(Request.Form("nik"),"")

	Nik=Nik.replace("/*","")
	Nik=Nik.replace("'","")
	Pass=Pass.replace("/*","")
	Pass=Pass.replace("'","")

	Records.Source="Select * from MEMBER where PSW='"+Pass+"' and NIK_NAME='"+Nik+"'"
	Records.Open()
	if (Records.EOF){ErrorMsg+="Неверный 'Пароль или псевдоним'.<br>"}
	else {if (Records("STATE").Value<=0){
		Session("id_mem")=String(Records("ID").Value)
		Session("name_mem")=String(Records("NAME").Value)
		Session("nik_mem")=String(Records("NIK_NAME").Value)
		Session("state_mem")=String(Records("STATE").Value)
		Session("is_adm_mem")=String(Records("ADM").Value)
		Session("email_mem")=String(Records("E_MAIL").Value)
		Session("phone_mem")=String(Records("PHONE").Value)
		Session("key_mem")=String(Records("KEYWORD").Value)
		} else {ErrorMsg+="Неверный 'Пароль или псевдоним'.<br>"}
	}
	Records.Close()
	if (ErrorMsg==""){
		if(Session("state_mem")==-1){Response.Redirect("activete_account.asp")} 
		Response.Redirect(Session("backurl"))
	}
}

%>

<html>
<head>
<title>Сибирский Аукцион Полезных Вещей</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<style type="text/css">
<!--
p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 10pt; font-weight: 400; color: #000000; margin:  3px}
a:hover {color:#FF3333; font: 9pt/9pt Arial, Helvetica, sans-serif}
a:visited {  font-family: Arial, Helvetica, sans-serif; color: #990000; font-size: 9pt; line-height: 9pt}
h1 {  font: 10pt/normal Verdana, Arial, Helvetica, sans-serif; color: #000000; text-decoration: blink; border-width: 1px auto medium medium; border-color: #CCCCFF black black; background: repeat-y; margin: 3px 3px}
a:link {  font: 9pt/9pt Arial, Helvetica, sans-serif; color: #000000}
-->
</style>
</head>
<body bgcolor="#FFFFFF" background="Images/block_back.gif">
<p><font face="Arial, Helvetica, sans-serif" size="2" color="#000099"><b> 
  <style>A:hover {
	COLOR: red
}
</style>
  </b></font></p>
<table align=center border=0 cellpadding=0 cellspacing=0 
width="100%" bgcolor="#3E0909">
  <tbody> 
  <tr> 
    <td width=150> 
      <h1 align=right><b><img src="Images/face.gif" width="91" height="60"></b></h1>
    </td>
    <td height=60 width=468> 
      <div align=center><font color=#006600 face="Arial, Helvetica, sans-serif" 
      size=6><bl><font color="#FFCC99" face="Times New Roman, Times, serif">СИБИРСКИЙ 
        АУКЦИОН</font><font color="#FFFFFF"><br>
        <font face="Arial, Helvetica, sans-serif" 
      size=6><b><font size="1" color="#CC9900">Аукцион Полезных Вещей </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <script language="">
var mydate=new Date()
var year=mydate.getYear()
if (year < 1000)
year+=1900
var day=mydate.getDay()
var month=mydate.getMonth()
var daym=mydate.getDate()
if (daym<10)
daym="0"+daym
var dayarray=new Array("Воскресенье","Понедельник","Вторник","Среда","Четверг","Пятница","Суббота")
var montharray=new Array("Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь")
document.write("<small><font color='#CC9900' face='Arial'><b>"+dayarray[day]+", "+montharray[month]+" "+daym+", "+year+"</b></font></small>")

</script>
        </font></b></font></font></bl></font></div>
    </td>
    <td>
      <!-- HotLog -->
      <script language="javascript">
hotlog_js="1.0";
hotlog_r=""+Math.random()+"&s=48805&im=15&r="+escape(document.referrer)+"&pg="+
escape(window.location.href);
document.cookie="hotlog=1; path=/"; hotlog_r+="&c="+(document.cookie?"Y":"N");
</script>
      <script language="javascript1.1">
hotlog_js="1.1";hotlog_r+="&j="+(navigator.javaEnabled()?"Y":"N")</script>
      <script language="javascript1.2">
hotlog_js="1.2";
hotlog_r+="&wh="+screen.width+'x'+screen.height+"&px="+
(((navigator.appName.substring(0,3)=="Mic"))?
screen.colorDepth:screen.pixelDepth)</script>
      <script language="javascript1.3">hotlog_js="1.3"</script>
      <script language="javascript">hotlog_r+="&js="+hotlog_js;
document.write("<a href='http://click.hotlog.ru/?48805' target='_top'><img "+
" src='http://hit3.hotlog.ru/cgi-bin/hotlog/count?"+
hotlog_r+"&' border=0 width=1 height=1></a>")</script>
      <noscript><a href=http://click.hotlog.ru/?48805 target=_top><img
src="http://hit3.hotlog.ru/cgi-bin/hotlog/count?s=48805&im=15" border=0 
width="1" height="1" ></a></noscript> 
      <!-- /HotLog -->
    </td>
  </tr>
  </tbody> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td>&nbsp; </td>
  </tr>
</table>
<table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td width=150> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan=2> 
      <p>:: <a href="auction.asp">Каталог</a> | <a href="reg_mem_auction.asp">Регистрация</a> 
        | <a href="auction_rules.html">Правила аукционов</a> | <a href="area.asp">Личный 
        кабинет</a> | <a href="reconnect.asp">Войти под другим именем</a> :: </p>
    </td>
  </tr>
  </tbody> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td>&nbsp; </td>
  </tr>
</table>
<table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td width=150> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan=2> 
      <h1><b>:: ВХОД ДЛЯ ЗАРЕГИСТРИРОВАННЫХ УЧАСТНИКОВ АУКЦИОНОВ ::</b></h1>
    </td>
  </tr>
  </tbody> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td>&nbsp; </td>
  </tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909">
  <tr valign="top"> 
    <td width="150" bordercolor="#FFFFFF" bgcolor="#FFCC99" background="Images/line_back.gif"> 
      <p>&nbsp;</p>
    </td>
    <td bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
        <p align="center"><%if(ErrorMsg!=""){%> </p>
  <center>
        <h2> 
          <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font><br>
            <%=ErrorMsg%></p>
        </h2>
  </center>
  <%}%> 
  <p>&nbsp;</p>
  <form name="form1" method="post" action="login.asp">
        <p align="center"><b>Если Вы участвуете в аукционах впервые, то зарегистрируйтесь 
          <a href="reg_mem_auction.asp">ЗДЕСЬ</a></b></p>
    <p>&nbsp;</p><table width="100%" border="0" cellspacing="4" cellpadding="0"> 
      <tr valign="middle"> 
        <td width="50%" align="right"> 
          <p><b>Псевдоним (логин)</b></p>
        </td>
        <td width="50%"> 
          <input type="text" name="nik" value="<%=isFirst?"":Request.Form("nik")%>" size="20" maxlength="20">
          * </td>
      </tr>
      <tr valign="middle"> 
        <td width="50%" align="right"> 
          <p><b>Пароль</b></p>
        </td>
        <td width="50%"> 
          <input type="password" name="pass" size="20" maxlength="20">
          * </td>
      </tr> 
    </table>
    <p align="center"> 
      <input type="submit" name="login" value="Ввод">
    </p>
    <hr size="1">
    <p align="center"><b>* - Обязательные поля</b></p>
    <p align="center">&nbsp;</p>
  </form>
    </td>
    <td width="35" bordercolor="#FFFFFF" height="408" bgcolor="#FFFFFF"> 
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td>&nbsp; </td>
  </tr>
</table>
<table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#3E0909" bordercolor="#3E0909"> 
    <td height=60 width=150> 
      <p>&nbsp;</p>
    </td>
    <td height=60 bgcolor="#3E0909"> 
      <script language="JavaScript" src="http://vbn.tyumen.ru/cgi-bin/hints.cgi?vbn&bloka">
</script>
    </td>
  </tr>
  </tbody> 
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909">
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td height="24"> 
      <h1 align="center"><font size="1">Программирование и дизайн <a href="http://www.rusintel.ru/" target="_blank">ЗАО 
        Русинтел</a> &copy; 2001-2002</font></h1>
    </td>
  </tr>
</table>
</body>
</html>
