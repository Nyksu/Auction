<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\creaters1.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->

<%
var name_lot=""
var tek_price=""
var valuta=""
var host_lot=""
var end_date=""
var name_img=""
var tip_nam=""
var tip=1
var kvo_memb=""
var subj_id=""
var kvo_stavok=0
var rate_state=""
var last_mem=""



if(String(Session("id_mem"))=="undefined"){
 Session("backurl")="area.asp"
 Response.Redirect("login.asp")
}

if(Session("state_mem")!=0){
 Session("backurl")="area.asp"
 Response.Redirect("activete_account.asp")
}

Records.Source="Select * from MEMBER where ID="+Session("id_mem")
   Records.Open()
   if (Records.EOF){
	Records.Close()
	Response.Redirect("auction.asp")
	}
	host_lot=Records("NAME").Value
   Records.Close()
   
   
%>

<html>
<head>
<title>Аукцион Полезных Вещей</title>
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

<meta http-equiv="refresh" content="45">
<link rel="stylesheet" href="/auction.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" background="/Images/block_back.gif">
<style>A:hover {
	COLOR: red
}
</style>

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
      size=6><font color=#FFFFFF face="Arial, Helvetica, sans-serif" 
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
        </font></b></font></font></font></font></font></bl></font></div>
    </td>
    <td> 
      <div> 
        <div align="right">
          <script language="JavaScript">
var code='';
</script>
          <script language="JavaScript" src="http://banners.isurgut.ru/GetBanner.asp?Member_id=296&Type_id=3"></script>
          <script language="JavaScript">
document.write(code);
</script>
        </div>
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
    <td colspan=2 bgcolor="#FFCC99"> 
      <h1><b>:: Личный кабинет :: </b><a href="support.asp">Поддержка</a> </h1>
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
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99" valign="top"> 
    <td width="150" bgcolor="#FFCC99" bordercolor="#CFE0FC"> </td>
    <td colspan=2 bgcolor="#FFCC99" bordercolor="#FFCC99"> 
      <h1><a href="lots_end_list.asp">Ваши лоты, торги на которые закончились</a><br>
        <b><a href="area.asp">Лоты, торги на которые продолжаются</a></b> </h1>
    </td>
  </tr>
  </tbody> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td><b><font color="#000066" face="Arial, Helvetica, sans-serif" size="3"><%=host_lot%></font></b> 
    </td>
  </tr>
</table>
<div align="right">
  <table width="100%" border="1" bordercolor="#330000">
    <tr bordercolor="#FFFFFF"> 
      <td bgcolor="#FFCC99" width="50%" bordercolor="#330000"> 
        <div align="center"><font size="3" color="#006600" face="Arial, Helvetica, sans-serif"><b><font size="2" color="#000066">-- 
          ВАШИ СТАВКИ (на закончивших торги аукционах)--</font></b></font></div>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF" valign="top"> 
      <td width="50%"> 
        <%
		sql=end_st_lots
		sql=sql.replace("%HOST_ID",Session("id_mem"))
		Records.Source=sql
		Records.Open()
		while (!Records.EOF){
			lot_id=String(Records("ID").Value)
			name_lot=String(Records("NAME").Value)
			tek_price=String(Records("PRICE").Value)
			valuta=String(Records("VALUTA").Value)
			tip_nam=String(Records("type_name").Value)
			tip=String(Records("AUCTION_TYPE_ID").Value)
			end_date=Records("DATE_END").Value
			name_subj=Records("SUBJNAME").Value
			subj_id=Records("AUCTION_SUBJ_ID").Value
			last_mem=Records("LAST_MEM").Value
   			if (tip==1) {nam_img="auc_regular"}
   			if (tip==2) {nam_img="auc_reserv"}
   			if (tip==4) {nam_img="auc_fix"}
   			if (tip==3) {nam_img="auc_inout"} 
			Records.MoveNext()
			sql="Select Count(*) as kvo from trade_history where goods_id="+lot_id
			Records1.Source=sql
			Records1.Open()
			kvo_stavok=Records1("KVO").Value
			Records1.Close()
			sql="Select * from GET_COUNT_LOT_MEMBERS("+lot_id+")"
			Records1.Source=sql
			Records1.Open()
			kvo_memb=Records1("KOLVO").Value
			Records1.Close()
			if (last_mem==Session("id_mem")) {rate_state="Играет Ваша ставка!!"} 
			else {rate_state="Ваша ставка перебита!! Далайте ставку!!"}
		%>
        <table width="100%" border="1" bordercolor="#330000">
          <tr bordercolor="#FFFFFF" bgcolor="#BDCECE"> 
            <td colspan="3"> 
              <p><font size="1"><b><img src="<%="Images/"+nam_img+".gif"%>" width="16" height="16" align="absmiddle" alt="<%=tip_nam%>"></b> 
                (лот № <%=lot_id%>)</font> <a href="auction.asp?subj_id=<%=subj_id%>&lot_id=<%=lot_id%>"><%=name_lot%></a> 
                <b><font size="2"> (<a href="auction.asp?subj_id=<%=subj_id%>"><%=name_subj%></a>)</font></b> </p>
            </td>
          </tr>
          <tr bordercolor="#FFFFFF" bgcolor="#FFFFFF"> 
            <td height="13" width="30%"> 
              <p>Кол-во участников: <%=kvo_memb%></p>
              <%=kvo_stavok==0?" Ставок нет!":" Принято ставок: "+kvo_stavok%></td>
            <td height="13" width="30%"> 
              <p>дата окончания: <font color="#3333FF"><%=end_date%></font></p>
            </td>
            <td height="13"> 
              <p><b>Текущая цена лота:</b><font color="#3333FF"> <b><font color="#FF0000"><%=tek_price%></font> 
                <%=valuta%> </b></font></p>
            </td>
          </tr>
          <tr bgcolor="#EBF3F2" bordercolor="#FFFFFF"> 
            <td colspan="3" height="25"> 
              <p><b><font color="#FF0000"><%=rate_state%></font></b></p>
            </td>
          </tr>
          <tr bgcolor="#BDCECE" bordercolor="#FFFFFF"> 
            <td colspan="3" height="16"> 
              <p align="center"><a href="do_rate.asp?subj=<%=subj_id%>&lot=<%=lot_id%>"><b>Сделать 
                ставку </b></a></p>
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
      <td height=60>
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
</div>

</body>
</html>
