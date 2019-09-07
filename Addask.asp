<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\email.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\next_id.inc" -->

<%
var subj=parseInt(Request("subj"))
var lot=parseInt(Request("lot"))
var ErrorMsg=""
var Text=""
var name=""
var tit=""
var uid=0
var name_lot=""
var host_lot=0
var tip=0
var nam_img=""
var hostname=""
var email=""
var id=0
var sql=""
var filename=""
var ts=""
var fs= new ActiveXObject("Scripting.FileSystemObject")


if (isNaN(subj)) {subj=0}
if (isNaN(lot)) {Response.Redirect("auction.asp")}
if (lot==0) {Response.Redirect("auction.asp?subj="+subj)}

if(String(Session("id_mem"))=="undefined"){
 Session("backurl")="addask.asp?subj="+subj+"&lot="+lot
 Response.Redirect("login.asp")
}

if(Session("state_mem")!=0){
 Session("backurl")="addask.asp?subj="+subj+"&lot="+lot
 Response.Redirect("activete_account.asp")
}

Records.Source="Select t1.*, t2.name as type_name from GOODS t1, auction_type t2 where t1.auction_type_id=t2.id and t1.ID="+lot
Records.Open()
if (Records.EOF){
		Records.Close()
		Response.Redirect("auction.asp?subj_id="+subj)
}
name_lot=String(Records("NAME").Value)
host_lot=Records("MEMBER_ID").Value
tip=Records("AUCTION_TYPE_ID").Value
if (tip==1) {nam_img="auc_regular"}
if (tip==2) {nam_img="auc_reserv"}
if (tip==4) {nam_img="auc_fix"}
if (tip==3) {nam_img="auc_inout"}   
Records.Close()

Records.Source="Select * from MEMBER where ID="+host_lot
Records.Open()
if (Records.EOF){
		Records.Close()
		Response.Redirect("auction.asp?subj_id="+subj)
}
hostname=Records("NIK_NAME").Value
email=TextFormData(Records("E_MAIL").Value,"")
Records.Close()

var eml=Server.CreateObject("JMail.Message")
var isSending=false

isFirst=String(Request.Form("Submit"))=="undefined"
ShowForm=true
if(!isFirst){
	
		//-------------input validation-----------
		tit=TextFormData(Request.Form("Name"),"")
		Text=TextFormData(Request.Form("txt"),"")

		if(Text.length>1000){ErrorMsg+="Сообщение превышает допустимый размер.<br>"}
		if(Text.length<4){ErrorMsg+="Сообщение отсутствует.<br>"}
		if(tit.length<3){ErrorMsg+="Укажите тему сообщения!<br>"}

		if (ErrorMsg==""){
			id=NextID("UNIVERSAL")
			filename=LotsFilePath+id+".ask"
			sql=asklotadd
			sql=sql.replace("%ID",id)
			sql=sql.replace("%NAME",tit)
			sql=sql.replace("%GD",lot)
			sql=sql.replace("%MEM",Session("id_mem"))
			Connect.BeginTrans()
			try{
				Connect.Execute(sql)
				ts=fs.OpenTextFile(filename,2,true)
				ts.Write(Text)
				ts.Close()
				Connect.CommitTrans()
				try {
				eml.Logging=false
				eml.From="auction@72rus.ru"
				eml.Subject="Задан вопрос по Вашему лоту!"
				eml.Charset=characterset
				eml.ContentTransferEncoding = "base64"
				eml.AddRecipient(email)
				eml.FromName="www.auction.72rus.ru"
				eml.AppendText("Через торговую систему задан вопрос по  Вашему лоту № : "+lot+"\n")
				eml.AppendText("Просим Вас по возможности ответить на вопросы на странице:\n")
				eml.AppendText("http://www.auction.72rus.ru/Asklst.asp?subj_id="+subj+"&lot_id="+lot+"\n\n")
			    eml.AppendText("Автоматическое сообщение. Сибирский аукцион: www.auction.72rus.ru")
				isSending=eml.Send(servsmtp)
				}
				catch(e){}
	 			if (!isSending) {ErrorMsg+="Проблемы с почтой.<br>"}
				ShowForm=false
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
			}
		} else { ShowForm=true }
}
%>

<html>
<head>
<title>Аукцион Нужных Вещей. Вопрос по лоту.</title>
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
      <h1><b>:: Вопрос по лоту № <%=lot%> ( <a href="auction.asp?subj_id=<%=subj%>&lot_id=<%=lot%>"><%=name_lot%></a> 
        ) ::</b></h1>
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
      <p><a href="auction_rules.html">Правила аукционов</a></p>
      <p><a href="area.asp">Личный кабинет</a></p>
      <p><a href="reconnect.asp">Войти под другим именем</a></p>
    </td>
    <td bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
<%if(ErrorMsg!=""){%>
      <center>
        <h2> 
          <p> <font color="#FF3300"><b>Ошибка!</b></font> <br>
            <%=ErrorMsg%></p>
        </h2>
      </center>
<%}%>
<%if(ShowForm){%>
      <form name="Guest" method="post" action="addask.asp">
        <p>
          <input type="hidden" name="subj" value="<%=subj%>">
          <input type="hidden" name="lot" value="<%=lot%>">
          <font color="#FF0000"><b>Все поля обязательны к заполнению</b></font> 
        </p>
        <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFCC99">
          <tr valign="middle" bgcolor="#EBF3F2" bordercolor="#FFCC99"> 
            <td width="200" align="right"> 
              <h1><b>Тема вопроса:</b></h1>
            </td>
            <td> 
              <input type="text" name="Name" value="<%=isFirst?"":Request.Form("Name")%>" size="50" maxlength="50">
            </td>
          </tr>
          <tr> 
            <td width="200" align="right" valign="top" bgcolor="#FFCC99"> 
              <h1><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Ваш 
                вопрос:</font></h1>
            </td>
            <td bgcolor="#FFCC99"> 
              <textarea name="txt" cols="50" rows="8"><%=Text%></textarea>
            </td>
          </tr>
        </table>
        <p align="center"> 
          <input type="submit" name="Submit" value="Сохранить вопрос">
        </p>
      </form>
      <%}
else{%>
      <center>
        <p><font size="3" color="#0000FF"><b>Вопрос размещен.</b></font></p>
        </center>
      <%}%>
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
