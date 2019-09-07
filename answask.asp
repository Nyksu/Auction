<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var subj=parseInt(Request("subj"))
var lot=parseInt(Request("lot"))
var name_lot=""
var askname=""
var filename=""
var ShowForm=true
var hostname=""
var txt=""
var tema=""
var ErrorMsg=""
var idask=parseInt(Request("ask"))
var filename=""
var ts=""
var host_lot=""
var sql=""

if (isNaN(subj)) {subj=0}
if (isNaN(lot)) {Response.Redirect("auction.asp")}
if (isNaN(idask)) {Response.Redirect("auction.asp")}

if(String(Session("id_mem"))=="undefined"){
 Session("backurl")="Addask.asp?subj="+subj+"&lot="+lot
 Response.Redirect("login.asp")
}

if(Session("state_mem")!=0){
 Session("backurl")="Addask.asp?subj="+subj+"&lot="+lot
 Response.Redirect("activete_account.asp")
}

if (lot>0){
   Records.Source="Select t1.*, t2.name as type_name from GOODS t1, auction_type t2 where t1.auction_type_id=t2.id and t1.ID="+lot
   Records.Open()
   if (Records.EOF){
		Records.Close()
		Response.Redirect("auction.asp?subj_id="+subj)
   }
   name_lot=String(Records("NAME").Value)
   host_lot=Records("MEMBER_ID").Value
   Records.Close()
   Records.Source="Select * from MEMBER where ID="+host_lot
   Records.Open()
   if (Records.EOF){
		Records.Close()
		Response.Redirect("auction.asp?subj_id="+subj)
   }
   hostname=Records("NIK_NAME").Value
   Records.Close()
 }
 else {Response.Redirect("auction.asp")}
 
if (host_lot!=Session("id_mem")){Response.Redirect("auction.asp?subj_id="+subj+"&lot_id="+lot)}

Records.Source= "Select * from ASK_GOODS where id="+idask
Records.Open()
   if (Records.EOF){
		Records.Close()
		Response.Redirect("auction.asp?subj_id="+subj)
   }
tema=String(Records("NAME").Value)
Records.Close()

Records.Source= "Select * from ANSWER where ask_goods_id="+idask
Records.Open()
   if (!Records.EOF){
		Records.Close()
		Response.Redirect("auction.asp?subj_id="+subj)
   }
Records.Close()
 
isFirst=String(Request.Form("Savebt"))=="undefined"

if (!isFirst) {
	var fs= new ActiveXObject("Scripting.FileSystemObject")
	var eml=Server.CreateObject("JMail.Message")
	txt=TextFormData(Request.Form("ask"),"")
	if (txt.length<3) {ErrorMsg+="Слишком короткий ответ! <br>"}
	
	if (ErrorMsg=="") {
		filename=LotsFilePath+idask+".ans"
		sql="Insert into answer (ask_goods_id, date_answ) values("+idask+", 'TODAY')"
		Connect.BeginTrans()
			try{
				Connect.Execute(sql)
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
			}
			if (ErrorMsg==""){
			 try
			 {
				ts=fs.OpenTextFile(filename,2,true)
				ts.Write(txt)
				ts.Close()
				Connect.CommitTrans()
				ShowForm=false
			 }
			 catch(e){
				ErrorMsg+=filename
				Connect.RollbackTrans()
				}
			}
	}
}
%>

<html>
<head>
<title>Добавление вопроса по лоту</title>
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

<body background="Images/block_back.gif">
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
      size=6><font color=#FFFFFF face="Arial, Helvetica, sans-serif" 
      size=6><b><font size="1" color="#CC9900">Аукцион Полезных Вещей / </font><font color=#FFFFFF face="Verdana, Arial, Helvetica, sans-serif" 
      size=2> 
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
        </font></b></font></font></font></bl></font></div>
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
<table border=1 bordercolor=#333333 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td width=150> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan=2> 
      <h1><b>:: Ответ на вопрос по лоту № <%=lot%> ( <a href="auction.asp?subj_id=<%=subj%>&lot_id=<%=lot%>"><%=name_lot%></a> 
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
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#333333">
  <tr valign="top"> 
    <td width="150" bordercolor="#FFCC99" bgcolor="#FFCC99" background="Images/line_back.gif"> 
      <p><a href="auction_rules.html">Правила аукционов</a></p>
      <p><a href="area.asp">Личный кабинет</a></p>
      <p><a href="reconnect.asp">Войти под другим именем</a></p>
    </td>
    <td bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
      <div align="right">
        <p align="center"> 
          <%if(ErrorMsg!=""){%>
		</p>
		<center>
		<h2> 
		<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
        <%=ErrorMsg%></p>
		</h2>
		</center>
		<%}%>
		</div>
	 <%if (ShowForm) {%>	
	<form name="form1" method="post" action="answask.asp">
	<input type="hidden" name="subj" value="<%=subj%>">
    <input type="hidden" name="lot" value="<%=lot%>">
        <input type="hidden" name="ask" value="<%=idask%>">
        <table width="100%" border="0" cellspacing="2" cellpadding="0">
        <tr bgcolor="#EBF3F2"> 
          <td height="2" bgcolor="#EBF3F2">
              <h1>Тема вопроса:<b><font color="#000000"> <%=tema%></font></b></h1>
            </td>
        
        <tr bgcolor="#EBF3F2"> 
            <td bgcolor="#FFCC99"> 
              <p> 
                <textarea name="ask" cols="50" rows="6"><%=txt%></textarea>
              </p>
            <p> 
                <input type="submit" name="Savebt" value="Сохранить ответ">
            </p>
          </td>
      </table>
	  </form>
	  <p>&nbsp;</p>
<%}else {%>
	<div align="center">
	    <p><font size="3" color="#0000FF"><b>Ответ размещен. </b></font></p>
	</div>
	<p align="center">&nbsp;</p>
	<p align="center"> 
<%}%>
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
