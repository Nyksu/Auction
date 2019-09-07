<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->

<%
var ErrorMsg=""
var sql=""
var subj=parseInt(Request("subj"))
var lot=parseInt(Request("lot"))
var subj_name=""
var res=""
var atensiontxt=""
var name_lot=""
var tek_price=0
var valuta=""
var kvo_stavok=0
var host_lot=""
var end_date=""
var start_date=""
var nam_img=""
var tip_nam=""
var stavka=0
var delta=0
var minrate=0
var atip=0
var start_price=0
var test_price=0


if (isNaN(subj)) {Response.Redirect("auction.asp")}
if (isNaN(lot)) {Response.Redirect("auction.asp?subj_id="+subj)}

if (subj>0){
   Records.Source="Select name from AUCTION_SUBJ where ID="+subj
   Records.Open()
   if (Records.EOF){
	Records.Close()
	Response.Redirect("auction.asp")
   }
   subj_name=String(Records("NAME").Value)
   Records.Close()
}

if(String(Session("id_mem"))=="undefined"){
 Session("backurl")="do_rate.asp?subj="+subj+"&lot="+lot
 Response.Redirect("login.asp")
}

if(Session("state_mem")!=0){
 Session("backurl")="do_rate.asp?subj="+subj+"&lot="+lot
 Response.Redirect("activete_account.asp")
}

if (lot>0){
   sql=select_allfor_lots
   sql=sql.replace("%LOT_ID",lot)
   Records.Source=sql
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
   delta=Records("DELTA").Value
   atip=Records("AUCTION_TYPE_ID").Value
   start_price=Records("START_PRICE").Value
   tip_nam=Records("type_name").Value
   tip=Records("AUCTION_TYPE_ID").Value
   if (tip==1) {nam_img="auc_regular"}
   if (tip==2) {nam_img="auc_reserv"}
   if (tip==4) {nam_img="auc_fix"}
   if (tip==3) {nam_img="auc_inout"}
   
   Records.Close()
   
   if (host_lot==Session("id_mem")){Response.Redirect("auction.asp?subj_id="+subj+"&lot_id="+lot)}
   
   minrate=start_price*delta/100
   if (atip<3) {
     atensiontxt="Вы можете добавить к текущей  стоимости сумму не мене "+minrate+" "+valuta
	 test_price=FloatFormData(tek_price)+FloatFormData(minrate)
	 }
   if (atip==3) {
     atensiontxt="Вы можете снизить текущую  стоимость на сумму не мене "+minrate+" "+valuta
	 test_price=FloatFormData(tek_price)-FloatFormData(minrate)
	 }
   if (atip==4) {
     atensiontxt="Для ставки по лоту с фиксированной ценой вы должны указать в графе 'Ваше предложение' текущую стоимость лота."
	 test_price=FloatFormData(tek_price)
   }
   

   Records.Source="Select * from MEMBER where ID="+host_lot
   Records.Open()
   if (Records.EOF){
	Records.Close()
	Response.Redirect("auction.asp?subj_id="+subj)
   }
   host_lot=Records("NIK_NAME").Value
   Records.Close()
   
   Records.Source="Select Count(id) as kvo from TRADE_HISTORY where GOODS_ID="+lot
   Records.Open()
   kvo_stavok=Records("KVO").Value
   Records.Close()
}

isFirst=String(Request.Form("Submit"))=="undefined"
ShowForm=true

if(!isFirst){
		stavka=TextFormData(Request.Form("stavka"),"")
		
		if (stavka==""){ErrorMsg+="Ставка не сделана!! <br>"}
		if ((atip<3)&&(stavka<test_price)){ErrorMsg+="Ставка меньше возможной!! Увеличте сумму предложения!! <br>"}
		if ((atip==3)&&(stavka>test_price)){ErrorMsg+="Ставка больше возможной!! Уменьшите сумму предложения!! <br>"}
		if ((atip==4)&&(stavka!=tek_price)){ErrorMsg+="Ставка не равна текущей стоимости!! Уравняйте сумму предложения с текущей стоимостью!! <br>"}
		

		if (ErrorMsg==""){
			sql=do_rate
			sql=sql.replace("%LOT",lot)
			sql=sql.replace("%RATE",stavka)
			sql=sql.replace("%MEMID",Session("id_mem"))
			Connect.BeginTrans()
			try{
				Records.Source=sql
				Records.Open()
   				res=Records("RES").Value
   				Records.Close()
				Connect.CommitTrans()
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
			}
			if (ErrorMsg==""){
			 	if (res==0) {ShowForm=false}
				else {
				    if (res==2) {ErrorMsg+="Лот закрыт для ставок!!! <br>"}
					if (res==1) {ErrorMsg+="Нарушение правил ставок!! <br>"}
				}
			 
			}
		}
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

<body bgcolor="#FFFFFF" background="/Images/block_back.gif">
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
        </font></b></font></font></font></font></font></font></bl></font></div>
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
    <td>&nbsp;</td>
  </tr>
</table>
<table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td width=150> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan=2> 
      <p>:: <a href="reg_mem_auction.asp">Регистрация</a> | <a href="auction_rules.html">Правила 
        аукционов</a> | <a href="area.asp">Личный кабинет</a> | <a href="reconnect.asp">Войти 
        под другим именем</a> :: </p>
    </td>
  </tr>
  </tbody> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td width=150> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan=2> 
      <h1><b>:: ВАША СТАВКА::</b></h1>
    </td>
  </tr>
  </tbody> 
</table>
<p align="center"> 
    <%if(ErrorMsg!=""){%>
  </p>
  <center>
  <h2> 
    <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> 
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
      <br>
      <%=ErrorMsg%></p>
  </h2>
  </center>
  <%}%>

<form name="form1" method="post" action="do_rate.asp">
  
    <input type="hidden" name="subj" value="<%=subj%>">
    <input type="hidden" name="lot" value="<%=lot%>">

    <%if (ShowForm) {%>
  <table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
    <tbody> 
    <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
      <td width=150> 
        <h1>&nbsp;</h1>
      </td>
      <td colspan=2> 
        <h1><b><a href="auction.asp">КАТАЛОГ</a>&gt;&gt; <a href="auction.asp?subj_id=<%=subj%>"><%=subj_name%></a> 
          &gt;&gt; <%=name_lot%></b></h1>
      </td>
    </tr>
    </tbody> 
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#3E0909" width="60">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#330000">
    <tr valign="top"> 
      <td width="150" bordercolor="#FFCC99" height="148" bgcolor="#FFCC99" background="Images/line_back.gif"> 
        <h1 align="right"><b>Описание&gt;&gt; </b></h1>
        <p>&nbsp;</p>
        </td>
      <td colspan="2" bgcolor="#FFFFFF" bordercolor="#FFFFFF" height="148"> 
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
          <tr> 
            <td bgcolor="#EBF3F2" colspan="2"> 
              <div align="center"> 
                <p align="left"><font size="1"><b><img src="<%="Images/"+nam_img+".gif"%>" width="16" height="16" align="absmiddle" alt="<%=tip_nam%>"></b> 
                  (лот № <%=lot%>)</font> <a href="auction.asp?subj_id=<%=subj%>&lot_id=<%=lot%>"><%=name_lot%></a></p>
              </div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFCC99" colspan="2" height="24"> 
              <p><b>Текущая цена лота:</b><font color="#3333FF"> <b><font color="#FF0000"><%=tek_price%></font> <%=valuta%></b></font><%=kvo_stavok==0?" Ставок нет!":" Уже сделано "+kvo_stavok+" ставок."%></p>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFCC" height="15" valign="top" colspan="2"> 
              <p><b><font color="#666666">Владелец лота</font>:</b> <%=host_lot%></p>
              <p>&nbsp;</p>
            </td>
          </tr>
          <tr> 
            <td> 
              <p>Начало торгов: <%=start_date%></p>
            </td>
            <td> 
              <p>Окончание торгов: <%=end_date%></p>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFCC"> 
              <p align="right"><b><font size="3" color="#FF0000">Внимание!!&nbsp;</font></b></p>
            </td>
            <td bgcolor="#FFFFCC"> 
              <p>&nbsp;<%=atensiontxt%></p>
            </td>
          </tr>
          <tr> 
            <td> 
              <p align="right"><b>Ваше предложение :</b></p>
            </td>
            <td> 
              <p align="left"> 
                <input type="text" name="stavka" value="<%=isFirst?test_price:Request.Form("stavka")%>">
                <font color="#3333FF"><b><%=valuta%> 
                <input type="submit" name="Submit" value="Сделать ставку">
                </b></font></p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <p align="center">&nbsp;</p>
  <p align="center">
    <%}else {%>
  </p>
  
</form>
<div align="right">
  <div align="center">
    <p><font size="3" color="#0000FF"><b>Ваша ставка принята. </b></font></p>
  </div>
  <p align="center">&nbsp;</p>
  <p align="center"> 
    <%}%>
  </p>
</div>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#990000">
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td width="150"> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan="2"> 
      <h1><a href="auction.asp"><b>КАТАЛОГ</b></a> <font face="Arial, Helvetica, sans-serif">&gt;&gt; 
        <font color="#6699FF"><a href="auction.asp?subj_id=<%=subj%>"><%=subj_name%></a></font></font> </h1>
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
</body>
</html>
