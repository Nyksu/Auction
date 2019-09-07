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
var newask=0

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
p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 10pt; font-weight: 400; color: #000000; margin:  3px}
a:hover {color:#FF3333; font-family: Arial, Helvetica, sans-serif}
a:visited {  font-family: Arial, Helvetica, sans-serif; color: #990000}
h1 {  font: 10pt/normal Verdana, Arial, Helvetica, sans-serif; color: #3E0909; text-decoration: blink; border-width: 1px auto medium medium; border-color: #CCCCFF black black; background: repeat-y; margin: 3px 3px}
a:link {  color: #000000; font-family: Arial, Helvetica, sans-serif}
-->
</style>
<meta http-equiv="refresh" content="300">
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
        <h1><b><font color="#FFCC99">Успешных торгов! <br>
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
    <td colspan=2> 
      <h1><b><font color="#000000">:: Личный кабинет :: </font></b><a href="support.asp">Поддержка</a> 
        :: <%if (Session("is_adm_mem")>0) {%><a href="adm.asp">Управление</a><%}%></h1>
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
    <td width="50%" bgcolor="#CFE0FC" bordercolor="#CFE0FC"> 
      <div align="center">
        <h1><a href="lots_end_list.asp">Ваши лоты, торги на которые закончились</a></h1>
      </div>
    </td>
    <td colspan=2 bgcolor="#F9F4DB" bordercolor="#F9F4DB"> 
      <h1 align="center"><a href="lots_win_list.asp">История Ваших выигравших 
        ставок</a></h1>
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

  <table width="100%" border="1" bordercolor="#3E0909">
    <tr bordercolor="#330000"> 
      <td bgcolor="#CFE0FC" width="50%" bordercolor="#330000"> 
        <div align="center">
          <h1><font size="2" face="Arial, Helvetica, sans-serif"><b>-- </b></font><font size="2"><b>ВАШИ 
            ДЕЙСТВУБЩИЕ ЛОТЫ</b></font><font size="2" face="Arial, Helvetica, sans-serif"><b> 
            --</b></font></h1>
        </div>
      </td>
      <td bgcolor="#F9F4DB" width="50%"> 
        <div align="center">
          <h1><font size="2" color="#000066" face="Arial, Helvetica, sans-serif"><b>-- 
            </b></font><font size="2"><b>ВАШИ ДЕЙСТВУЮЩИЕ СТАВКИ</b></font><font size="2" color="#000066" face="Arial, Helvetica, sans-serif"><b> 
            --</b></font></h1>
        </div>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF" valign="top" bgcolor="#FFFFFF"> 
      <td width="50%" bordercolor="#FFFFFF">
<%
		sql=select_own_lots
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
			sql="Select Count(*) as kvo from ask_goods t1 where not exists (Select s1.* from answer s1 where s1.ask_goods_id=t1.id)  and t1.goods_id="+lot_id
			Records1.Source=sql
			Records1.Open()
			newask=Records1("KVO").Value
			Records1.Close()
		%>
        <table width="100%" border="1" bordercolor="#3E0909">
          <tr bordercolor="#FFFFFF" bgcolor="#CFE0FC"> 
            
          <td colspan="2"> 
            <p><font size="1">(лот № <%=lot_id%>)</font> <font size="3"><b><a href="auction.asp?subj_id=<%=subj_id%>&lot_id=<%=lot_id%>"><%=name_lot%></a> <img src="<%="Images/"+nam_img+".gif"%>" align="absmiddle" alt="<%=tip_nam%>" width="65" height="18"></b></font></p>
            </td>
          </tr>
          <tr bordercolor="#FFFFFF"> 
            
          <td colspan="2" height="23"> 
            <p><font size="2">Тема <a href="auction.asp?subj_id=<%=subj_id%>"><%=name_subj%></a></font> <b> </b> <%=kvo_stavok==0?" Ставок нет!":" Принято ставок: "+kvo_stavok%> </p>
          </td>
          </tr>
          <tr bgcolor="#EBF3F2" bordercolor="#FFFFFF"> 
            <td width="50%"> 
              <p>Кол-во участников: <%=kvo_memb%></p>
            </td>
            <td width="50%"> 
              <p>дата окончания: <font color="#3333FF"><%=end_date%></font></p>
            </td>
          </tr>
          <tr bgcolor="#EBF3F2" bordercolor="#FFFFFF"> 
            <td width="50%"> 
              <p><b>Текущая цена лота:</b><font color="#3333FF"> <b><font color="#FF0000"><%=tek_price%></font> 
                <%=valuta%></b></font></p>
            </td>
            <td width="50%" bgcolor="#EBF3F2"> 
              
            <p>Новых <a href="Asklst.asp?subj_id=<%=subj_id%>&lot_id=<%=lot_id%>">вопросов</a> 
              по лоту: <font color="#FF0000"><b><%=newask%></b></font></p>
            </td>
          </tr>
          <tr bgcolor="#EBF3F2" bordercolor="#FFFFFF"> 
            <td colspan="2" height="16"> 
              <p align="center"><a href="edit_lot.asp?subj=<%=subj_id%>&lot=<%=lot_id%>">Изменить 
                </a> | <a href="del_lot.asp?subj=<%=subj_id%>&lot=<%=lot_id%>"> 
                Удалить </a> | <a href="prolong_lot.asp?subj=<%=subj_id%>&lot=<%=lot_id%>"> 
                Продлить</a> | <a href="stop_lot.asp?subj=<%=subj_id%>&lot=<%=lot_id%>"> 
                Остановить</a></p>
            </td>
          </tr>
        </table>
        <p>&nbsp;</p>
        <p>
          <%} Records.Close()%>
        </p>
      </td>
      <td width="50%"> 
        <%
		sql=select_st_lots
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
        <table width="100%" border="1" bordercolor="#3E0909">
          <tr bordercolor="#FFFFFF" bgcolor="#F9F4DB"> 
            
          <td colspan="2"> 
            <p><font size="1">(лот № <%=lot_id%>)</font> <font size="3"><a href="auction.asp?subj_id=<%=subj_id%>&lot_id=<%=lot_id%>"><b><%=name_lot%></b></a></font> <img src="<%="Images/"+nam_img+".gif"%>" align="absmiddle" alt="<%=tip_nam%>" width="65" height="18"></p>
            </td>
          </tr>
          <tr bordercolor="#FFFFFF"> 
            
          <td colspan="2" height="13"> 
            <p><font size="2">Тема <a href="auction.asp?subj_id=<%=subj_id%>"><%=name_subj%></a></font> <b> </b> <%=kvo_stavok==0?" Ставок нет!":" Принято ставок: "+kvo_stavok%> </p>
          </td>
          </tr>
          <tr bgcolor="#EBF3F2" bordercolor="#FFFFFF"> 
            <td width="50%"> 
              <p>Кол-во участников: <%=kvo_memb%></p>
            </td>
            <td width="50%"> 
              <p>дата окончания: <font color="#3333FF"><%=end_date%></font></p>
            </td>
          </tr>
          <tr bgcolor="#EBF3F2" bordercolor="#FFFFFF"> 
            <td width="50%" bgcolor="#EBF3F2"> 
              <p><b>Текущая цена лота:</b><font color="#3333FF"> <b><font color="#FF0000"><%=tek_price%></font> 
                <%=valuta%> </b></font></p>
            </td>
            <td width="50%"> 
              <p><b><font color="#FF0000"><%=rate_state%></font></b></p>
            </td>
          </tr>
          <tr bgcolor="#EBF3F2" bordercolor="#FFFFFF"> 
            <td colspan="2" height="16"> 
              <p align="center"><a href="do_rate.asp?subj=<%=subj_id%>&lot=<%=lot_id%>">Сделать 
                ставку</a></p>
            </td>
          </tr>
        </table>
        <p>&nbsp;</p>
        <p>
          <%} Records.Close()%>
        </p>
      </td>
    </tr>
  </table>
  
  
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td><b><font color="#FFCC99"> 
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
      </font></b> </td>
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
