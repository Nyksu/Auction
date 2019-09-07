<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\creaters1.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var subj=parseInt(Request("subj_id"))
var lot=parseInt(Request("lot_id"))
var name_lot=""
var askname=""
var aunik=""
var asrid=0
var filename=""
var filname=""
var askdate=""
var fs= new ActiveXObject("Scripting.FileSystemObject")

if (isNaN(subj)) {subj=0}
if (isNaN(lot)) {Response.Redirect("auction.asp")}

if (lot>0){
   Records.Source="Select t1.*, t2.name as type_name from GOODS t1, auction_type t2 where t1.auction_type_id=t2.id and t1.ID="+lot
   Records.Open()
   if (Records.EOF){
	Records.Close()
	Response.Redirect("auction.asp?subj_id="+subj)
   }
   name_lot=String(Records("NAME").Value)
   tek_price=String(Records("PRICE").Value)
   valuta=String(Records("VALUTA").Value)
   start_date=Records("DATE_BEGIN").Value
   end_date=Records("DATE_END").Value
   host_lot=Records("MEMBER_ID").Value
   tip_nam=Records("type_name").Value
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
   host_lot=Records("NIK_NAME").Value
   Records.Close()
   

}

%>

<html>
<head>
<title>Вопросы по лоту № <%=lot%> ( <%=name_lot%> )</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<style type="text/css">
<!--
p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 10pt; font-weight: 400; color: #333333; margin:  3px}
a:hover {color:#FF3333}
a:visited {  font-family: Arial, Helvetica, sans-serif; color: #990000}
h1 {  font: 10pt/normal Verdana, Arial, Helvetica, sans-serif; color: #990000; border-width: 1px auto medium medium; border-color: #CCCCFF black black; background: repeat-y; margin: 3px 3px}
a:link {  color: #333333; font-family: Arial, Helvetica, sans-serif}
-->
</style>
</head>
<body bgcolor="#FFFFFF" background="/Images/block_back.gif">
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
      size=6><font color=#006600 face="Arial, Helvetica, sans-serif" 
      size=6><font color="#FFFFFF"><font face="Arial, Helvetica, sans-serif" 
      size=6><font color=#006600 face="Arial, Helvetica, sans-serif" 
      size=6><font color="#FFFFFF"><font face="Arial, Helvetica, sans-serif" 
      size=6><font face="Arial, Helvetica, sans-serif" 
      size=6><b><font size="1" color="#CC9900">Аукцион Полезных Вещей / </font><font color=#006600 face="Arial, Helvetica, sans-serif" 
      size=6><font color="#FFFFFF"><font face="Arial, Helvetica, sans-serif" 
      size=6><font face="Arial, Helvetica, sans-serif" 
      size=6><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b> 
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
        </b></font></b></font></font></font></font></b></font></font></font></font><font face="Arial, Helvetica, sans-serif" 
      size=6><b></b></font></font></font></font><font face="Arial, Helvetica, sans-serif" 
      size=6><b></b></font></font></font></bl></font></div>
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
    <td> 
      <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b> 
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
document.write("<small><font color='#006666' face='Arial'><b>"+dayarray[day]+", "+montharray[month]+" "+daym+", "+year+"</b></font></small>")

</script>
        </b></font> </div>
    </td>
  </tr>
</table>
<table border=1 bordercolor=#333333 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td width=150> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan=2> 
      <p><a href="auction.asp">Каталог</a> | <a href="reg_mem_auction.asp"></a><a href="reg_mem_auction.asp">Регистрация</a> 
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
      <h1><b>:: Вопросы по лоту № <%=lot%> ( <a href="auction.asp?subj_id=<%=subj%>&lot_id=<%=lot%>"><%=name_lot%></a> ) :: &lt;&lt; <a href="Addask.asp?lot=<%=lot%>&subj=<%=subj%>">Задать 
        вопрос по лоту</a></b></h1>
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
<%
Records.Source="Select t1.*, t2.NIK_NAME from ASK_GOODS t1, MEMBER t2 where t1.MEMBER_ID=t2.ID and t1.GOODS_ID="+lot+" order by t1.ID DESC"
Records.Open()
while (!Records.EOF)
{
	askid=String(Records("ID").Value)
	askname=String(Records("NAME").Value)
	aunik=String(Records("NIK_NAME").Value)
	askdate=Records("ASK_DATE").Value
	Records.MoveNext()
	 filename=LotsFilePath+askid+".ask"
     if( ! fs.FileExists(filename)) {filename=""}
     if (filename != "") {ts= fs.OpenTextFile(filename)}
%>
      <table width="100%" border="0" cellspacing="2" cellpadding="0">
        <tr bgcolor="#FFCC99"> 
          <td height="7"> 
            <div align="center"> 
              <h1 align="left"><b><font color="#000000">Тема вопроса: <font color="#000099"><%=askname%></font> 
                От <font color="#000099"><%=askdate%></font> </font></b></h1>
              <p align="left"><b><font color="#000000">Автор: <font color="#000099"><%=aunik%></font></font></b></p>
            </div>
          </td>
          <td colspan="2" height="7" width="141"> 
            <div align="center"> 
              <h1><b><font color="#333333">Управление:</font></b></h1>
            </div>
          </td>
        </tr>
        <tr bgcolor="#EBF3F2"> 
          <td height="2"> 
<%
	if (filename != "") {		
		while (!ts.AtEndOfStream){
		str=ts.ReadLine()
		while(str.indexOf("<")>=0){str=str.replace("<","&lt;")}		
		Response.Write(str+"<br>")}
 		ts.Close()
	}
%>
          </td>
          <td colspan="2" height="2" width="141"> 
            <p align="center"><a href="answask.asp?lot=<%=lot%>&subj=<%=subj%>&ask=<%=askid%>">ответить</a></p>
          </td>
		  </tr>
		  <tr bgcolor="#EBF3F2"> 
          <td height="2" bgcolor="#FFFFCC"> 
            <h1>Ответ:  </h1>
			<p>
              <%
	 filname=LotsFilePath+askid+".ans"
     if( ! fs.FileExists(filname)) {filname=""}
     if (filname != "") {ts= fs.OpenTextFile(filname)	
		while (!ts.AtEndOfStream){
		str=ts.ReadLine()
		while(str.indexOf("<")>=0){str=str.replace("<","&lt;")}		
		Response.Write(str+"<br>")}
 		ts.Close()
	}
%>   
          </p>
          </td>
          <td colspan="2" height="2" width="141" bgcolor="#FFFFCC"> 
            <p align="center">&nbsp;</p>
          </td>
		  </tr>
     </table>
<%} Records.Close()%>
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
