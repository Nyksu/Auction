<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
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
var nik=""
var lot_id=""
var fio=""
var city=""
var email=""
var memid=0
var regdate=""
var st=""
var ix=0


if(String(Session("id_mem"))=="undefined"){
 Session("backurl")="adm.asp"
 Response.Redirect("login.asp")
}

if(Session("state_mem")!=0){
 Session("backurl")="adm.asp"
 Response.Redirect("activete_account.asp")
}

if (Session("is_adm_mem")==0) { Response.Redirect("area.asp")}

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
p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 10pt; font-weight: 400; color: #000000; margin:  3px}
a:hover {color:#FF3333; font-family: Arial, Helvetica, sans-serif}
a:visited {  font-family: Arial, Helvetica, sans-serif; color: #990000}
h1 {  font: 10pt/normal Verdana, Arial, Helvetica, sans-serif; color: #3E0909; text-decoration: blink; border-width: 1px auto medium medium; border-color: #CCCCFF black black; background: repeat-y; margin: 3px 3px}
a:link {  color: #000000; font-family: Arial, Helvetica, sans-serif}
-->
</style>
</head>
<body bgcolor="#FFFFFF" background="Images/block_back.gif">
 
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
      size=6 color="#FFFFFF"><b><font size="1" color="#CC9900">Аукцион Полезных 
        Вещей / </font><font color=#FFFFFF face="Verdana, Arial, Helvetica, sans-serif" 
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
        </font></b></font></font></bl></font></div>
    </td>
    <td>
      <div align="center"> 
        <h1><b><font color="#FFCC99">Привет</font><font color="#FFCC99">! <br>
          <%=host_lot%></font></b> </h1>
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
      <h1><b><font color="#000000">:: Управление аукционом:: Все пользователи 
        :: Все лоты ::</font></b></h1>
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

<table width="100%" border="0">
  <tr bordercolor="#330000"> 
    <td width="50%" bordercolor="#330000" valign="top"> 
      <table width="100%" border="1">
        <tr bgcolor="#CFE0FC" bordercolor="#330000"> 
          <td colspan="2"> 
            <div align="center"> 
              <h1><font size="2"><b>-- Новые аукционы (за 3 дня) --</b></font></h1>
            </div>
          </td>
        </tr>
<%
sql=select_new_goods
Records.Source=sql
Records.Open()
while (!Records.EOF)
{
	lot_id=String(Records("ID").Value)
	name_lot=String(Records("NAME").Value)
	tek_price=String(Records("PRICE").Value)
	valuta=String(Records("VALUTA").Value)
	tip_nam=String(Records("type_name").Value)
	tip=String(Records("AUCTION_TYPE_ID").Value)
	nik=String(Records("NIK_NAME").Value)
	subj=Records("AUCTION_SUBJ_ID").Value
	regdate=Records("date_begin").Value
   	if (tip==1) {nam_img="auc_regular"}
   	if (tip==2) {nam_img="auc_reserv"}
   	if (tip==4) {nam_img="auc_fix"}
   	if (tip==3) {nam_img="auc_inout"} 
	Records.MoveNext()
%>
        <tr> 
          <td bgcolor="#FFFFFF" bordercolor="#330000"> 
            <p><font size="1">(лот № <%=lot_id%>)</font> <a href="auction.asp?subj_id=<%=subj%>&lot_id=<%=lot_id%>"><%=name_lot%></a> 
              (<%=nik%>) <font color="#3333FF"><b><%=tek_price%></b></font> <b><font color="#3E0909"><%=valuta%> 
              | </font></b><font color="#3E0909" size="1">добавлен: <%=regdate%></font></p>
          </td>
          <td bgcolor="#FFFFFF" bordercolor="#330000" width="67"><img src="<%="Images/"+nam_img+".gif"%>" width="65" height="18" align="absmiddle" alt="<%=tip_nam%>"></td>
        </tr>
        <%} Records.Close()%>
      </table>
    </td>
    <td valign="top" width="50%" bordercolor="#330000"> 
      <table width="100%" border="1">
        <tr> 
          <td bgcolor="#F9F4DB" bordercolor="#330000"> 
            <div align="center"> 
              <h1><font size="2" color="#000066"><b>-- Все пользователи --</b></font></h1>
            </div>
          </td>
        </tr>
<%
sql="Select * from member order by reg_date desc"
//sql="Select * from member where reg_date='TODAY' order by name"
Records.Source=sql
Records.Open()
ix=0
while (!Records.EOF)
{
	memid=String(Records("ID").Value)
	fio=String(Records("NAME").Value)
	city=String(Records("CITY").Value)
	email=String(Records("E_MAIL").Value)
	nik=String(Records("NIK_NAME").Value)
	regdate=Records("REG_DATE").Value
	st=Records("STATE").Value
	ix=ix+1
	Records.MoveNext()
%>
        <tr> 
          <td bgcolor="#FFFFFF" bordercolor="#330000"> 
            <p><%=ix%>. <a href="member.asp?mem=<%=memid%>"><%=nik%></a> (<font color="#FF0000" size="1"><%=fio%></font>) 
              <font color="#0000FF"> г. <%=city%></font>, <%=email%> | <font color="#FF0000"><%=regdate%></font> 
              | <%=st%></p>
          </td>
        </tr>
<%} Records.Close()%>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td><b><font color="#FFCC99"> </font></b> </td>
  </tr>
</table>
<table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#3E0909" bordercolor="#3E0909"> 
    <td height=60 width=150> 
      <p><b> 
        <script language="JavaScript">
var code='';
</script>
        <script language="JavaScript" src="http://banners.isurgut.ru/GetBanner.asp?Member_id=296&Type_id=3"></script>
        <script language="JavaScript">
document.write(code);
</script>
        </b></p>
    </td>
    <td height=60 bgcolor="#3E0909" bordercolor="#3E0909"> 
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
