<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\email.inc" -->

<%
var ErrorMsg=""
var text=""
var name=""
var tit=""
var uid=0
var Text=""

if(String(Session("id_mem"))=="undefined"){
 Session("backurl")="support.asp"
 Response.Redirect("login.asp")
}

if(Session("state_mem")!=0){
 Session("backurl")="support.asp"
 Response.Redirect("activete_account.asp")
}

var eml=Server.CreateObject("JMail.Message")
eml.Logging=false
eml.From=fromaddres
eml.AddRecipient(Recipient)
eml.Charset=characterset
eml.ContentTransferEncoding = "base64"

var isSending=false

isFirst=String(Request.Form("Submit"))=="undefined"
ShowForm=true
if(!isFirst){
	
		//-------------input validation-----------
		tit=TextFormData(Request.Form("Name"),"")
		Text=TextFormData(Request.Form("txt"),"")


		if(Text.length>2000){ErrorMsg+="Сообщение превышает допустимый размер.<br>"}
		if(Text.length<4){ErrorMsg+="Сообщение отсутствует.<br>"}
		if(tit.length<3){ErrorMsg+="Укажите тему сообщения!<br>"}
		

		if (ErrorMsg==""){
			
			try{
				eml.FromName=Session("nik_mem")
				eml.Subject=tit+" (Auction.72RUS.ru)"
				eml.AppendText(" Поддержка  Аукциона 72RUS \n")
				eml.AppendText(" Ф.И.О. : "+Session("name_mem")+"\n")
				eml.AppendText(" Email : "+Session("email_mem")+"\n")
				eml.AppendText("\n Сообщение : \n \n")
				eml.AppendText(TextFormData(Request.Form("txt")))
				isSending=eml.Send(servsmtp)
	 			if (isSending) {ShowForm=false}
			}
			catch(e){
				ErrorMsg+="Проблемы с почтой.<br>"
			}
		} else { ShowForm=true }
}
%>

<html>
<head>
<title>Аукцион Нужных Вещей</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<style type="text/css">
<!--
p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 10pt; font-weight: 400; color: #333333; margin:  3px}
a:hover {color:#FF3333}
a:visited {  color: #333333}
h1 {  font: 10pt/normal Verdana, Arial, Helvetica, sans-serif; color: #990000; text-decoration: blink; border-width: 1px auto medium medium; border-color: #CCCCFF black black; background: repeat-y; margin: 3px 3px}
a:link {  color: #990000}
-->
</style>
</head>
<body bgcolor="#FFFFFF" background="Images/block_back.gif">
<p><font face="Arial, Helvetica, sans-serif"><font size="2"><b><font face="Arial, Helvetica, sans-serif"><font size="2"><b><font face="Arial, Helvetica, sans-serif"><font size=2><b><font 
face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"><font size=2><b><font 
face="Arial, Helvetica, sans-serif"><font size=2><b><font color=#000099>
  <style>A:hover {
	COLOR: red
}
</style>
  </font></b></font></font></b></font></font></font></b></font></font></b></font></font></b></font></font></p>
<table align=center border=0 cellpadding=0 cellspacing=0 
width="100%" bgcolor="#3E0909">
  <tbody> 
  <tr> 
    <td width=150> 
      <h1 align=right><b></b><b><img src="Images/face.gif" width="91" height="60"></b></h1>
    </td>
    <td height=60 width=468> 
      <div align=center><font color=#006600 face="Arial, Helvetica, sans-serif" 
      size=6><bl><font color="#FFCC99" face="Times New Roman, Times, serif">СИБИРСКИЙ 
        АУКЦИОН</font><font color="#FFFFFF"><br>
        <font face="Arial, Helvetica, sans-serif" 
      size=6><font face="Arial, Helvetica, sans-serif" 
      size=6><b><font size="1" color="#CC9900">Аукцион Полезных Вещей </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>
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
        </b></font></b></font></font></font></bl></font></div>
    </td>
    <td> 
      <div align="right">
        <script language="JavaScript">
var code='';
</script>
        <script language="JavaScript" src="http://banners.isurgut.ru/GetBanner.asp?Member_id=296&Type_id=3"></script>
        <script language="JavaScript">
document.write(code);
</script>
      </div>
    </td>
  </tr>
  </tbody> 
</table>
<div align="left"></div>
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
      <h1><b>:: SUPPORT ДЛЯ ЗАРЕГИСТРИРОВАННЫХ УЧАСТНИКОВ АУКЦИОНОВ ::</b></h1>
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
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td valign="top"> 
            <table width="100%" border="0" cellspacing="6" cellpadding="0" bordercolor="#CCCCFF" align="center">
              <tr bordercolor="#CCCCFF" valign="middle" align="center"> 
                <td bordercolor="#CCCCFF" height="11"> 
                  <h1 align="left"><b>Служба поддержки</b></h1>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <%if(ErrorMsg!=""){%>
      <center>
        <h2> 
          <p> <font color="#FF3300">Ошибка!</font> <br>
            <%=ErrorMsg%></p>
        </h2>
      </center>
      <%}%>
      <form name="Guest" method="post" action="support.asp">
        <%if(ShowForm){%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr valign="middle"> 
            <td width="200" align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Тема 
              вопроса: </font></td>
            <td> 
              <input type="text" name="Name" value="<%=isFirst?"":Request.Form("Name")%>" size="50" maxlength="50">
            </td>
          </tr>
          <tr> 
            <td width="200" align="right" valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Ваш 
              вопрос:</font></td>
            <td> 
              <textarea name="txt" cols="50" rows="8"><%=text%></textarea>
            </td>
          </tr>
        </table>
        <p align="center"> 
          <input type="submit" name="Submit" value="Переслать">
          <input type="reset" name="Submit2" value="Очистить">
        </p>
        <%}
else{%>
        <center>
          <h2>Спасибо!<br>
            Ваше сообщение отправлено.</h2>
        </center>
        <%}%>
        <hr noshade size="1">
      </form>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="1"> </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
      </font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font> 
      <p align="center">&nbsp;</p>
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
      <p>
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
      </p>
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
