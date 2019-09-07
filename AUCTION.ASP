<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\creaters1.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->

<%
var subj=parseInt(Request("subj_id"))
var lot=parseInt(Request("lot_id"))
var name_subj=""
var name_lot=""
var tek_price=""
var valuta=""
var lot_about=""
var start_date=""
var end_date=""
var host_lot=""
var name_img=""
var sql=""
var sb_id=0
var lot_id=0
var kvo_stavok=0
var nam_img=""
var tip_nam=""
var newask=0
var tip=1
var image_name="lots/noimg.jpg"
var fs= new ActiveXObject("Scripting.FileSystemObject")
var tit=" > "
var city=""
var kask=0

if (isNaN(subj)) {subj=0}
if (isNaN(lot)) {lot=0}

if (subj>0){
   Records.Source="Select name from AUCTION_SUBJ where ID="+subj
   Records.Open()
   if (Records.EOF){
	Records.Close()
	Response.Redirect("auction.asp")
   }
   name_subj=String(Records("NAME").Value)
   tit=tit+name_subj
   Records.Close()
}

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
   tit=tit+" > "+name_lot
   Records.Source="Select * from MEMBER where ID="+host_lot
   Records.Open()
   if (Records.EOF){
	Records.Close()
	Response.Redirect("auction.asp?subj_id="+subj)
   }
   host_lot=Records("NIK_NAME").Value
   city=Records("CITY").Value
   Records.Close()
   
   Records.Source="Select Count(id) as kvo from TRADE_HISTORY where GOODS_ID="+lot
   Records.Open()
   kvo_stavok=Records("KVO").Value
   Records.Close()
   
	Records.Source="Select Count(*) as kvo from ask_goods t1 where not exists (Select s1.* from answer s1 where s1.ask_goods_id=t1.id)  and t1.goods_id="+lot
	Records.Open()
	newask=Records("KVO").Value
	Records.Close()
	
	Records.Source="Select Count(*) as kvo from ask_goods where goods_id="+lot
	Records.Open()
	kask=Records("KVO").Value
	Records.Close()
   
   filename=LotsFilePath+lot+LotsExtention
	if( ! fs.FileExists(filename)) {filename=""}
   if(filename!="") {ts= fs.OpenTextFile(filename)}
   if(fs.FileExists(LotsFilePath+lot+".gif")){image_name=GlobLotsFilePath+lot+".gif"}
   if(fs.FileExists(LotsFilePath+lot+".jpg")){image_name=GlobLotsFilePath+lot+".jpg"}
}

%>

<html>
<head>
<title><%=tit%> << Сибирский Аукцион Полезных Вещей</title>
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
<body bgcolor="#74BADC">
<p>
<style>A:hover {
	COLOR: red
}
</style>
<table align=center border=1 cellpadding=0 cellspacing=0 
width="100%" bgcolor="#CDE0DE" bordercolor="#330000">
  <tbody> 
  <tr bordercolor="#CDE0DE"> 
    <td width=150 align="center"> 
      <h1><img src="Images/Auction.gif" width="60" height="60"></h1>
    </td>
    <td height=60 align="center"><font face="Arial, Helvetica, sans-serif" 
      size=6><font face="Times New Roman, Times, serif" size="4">СИБИРСКИЙ АУКЦИОН</font></font> 
    </td>
    <td align="right" width="468"> 
      <script language="javascript" src="http://72rus.ru/banshow.asp?rid=9"></script></td>
  </tr>
  </tbody> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="150" align="center"><img src="Images/arrow.gif" width="20" height="10"></td>
    <td> 
      <p><marquee align="middle"
 height="10" scrolldelay="0"><font color="#330000" face="Verdana, Arial, Helvetica, sans-serif" size="1">Приветствуем 
        Вас в &quot;Торговой системе&quot;СИБИРСКИЙ АУКЦИОН&quot;! ................................ 
        Для того, чтобы принять участие в торгах или выставить свой собственный 
        лот на продажу - каждый желающий может пройти несложную регистрацию, получить 
        собственный пароль для работы в системе. Не забудьте ознакомиться с <a href="auction_rules.html">Правилами 
        аукционов</a> Желаем успешных торгов!</font></marquee></p>
    </td>
  </tr>
</table>
<%if (subj==0) {%>
<table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#FFEFDF" bordercolor="#FFEFDF"> 
    <td width=150 bordercolor="#CDE0DE" bgcolor="#CDE0DE"> 
      <h1><b>auction.72rus.ru</b></h1>
    </td>
    <td colspan=2 bordercolor="#CDE0DE" bgcolor="#CDE0DE"> 
      <p><a href="reg_mem_auction.asp"> Регистрация</a> | <a href="auction_rules.html">Правила 
        аукционов</a> | <a href="area.asp">Личный кабинет</a> | <a href="reconnect.asp">Войти 
        под другим именем</a> :: </p>
    </td>
  </tr>
  </tbody> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="150" align="center"><img src="Images/arrow.gif" width="20" height="10"></td>
    <td>&nbsp; </td>
  </tr>
</table>
<table border="1" cellspacing="2" cellpadding="0" bordercolor="#3E0909" align="left" width="150" bgcolor="#FFFFFF">
  <tr bordercolor="#330000"> 
    <td bgcolor="#CDE0DE" bordercolor="#330000"> 
      <div align="CENTER"> 
        <h1><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Ваши 
          действия </font></b></h1>
      </div>
    </td>
  </tr>
  <tr valign="top"> 
    <td bordercolor="#CDE0DE" bgcolor="#FFFFFF"> 
      <p><a href="auction_rules.html">Правила аукционов</a></p>
      <p><a href="area.asp">Личный кабинет</a></p>
      <p><a href="reconnect.asp">Войти под другим именем</a> </p>
    </td>
  </tr>
</table>
<p><img src="Images/arrow_rt.gif" width="10" height="20"></p>
<table width="126" border="1" cellspacing="2" cellpadding="0" align="right" bordercolor="#330000" bgcolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CDE0DE"> 
      <div align="CENTER"> 
        <h1><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Реклама</font></b></h1>
      </div>
    </td>
  </tr>
  <tr> 
    <td bordercolor="#CDE0DE">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <table border="1" cellspacing="2" cellpadding="0" width="100%" bordercolor="#3E0909" bgcolor="#FFFFFF" align="center">
        <tr bgcolor="#CDE0DE"> 
          <td height="2" bgcolor="#CDE0DE"> 
            <div align="center"> 
              <h1 align="left"><b>/ Аукционы</b></h1>
            </div>
          </td>
          <td colspan="2" height="2" width="141"> 
            <div align="center"> 
              <h1><b>Лоты</b></h1>
            </div>
          </td>
          <%
Records.Source="Select id_subj as ID, subj as NAME, coun_auc as KVO from GET_SUBJ_COUNT_LIST order by SUBJ"
Records.Open()
while (!Records.EOF)
{
	sb_id=String(Records("ID").Value)
	name_subj=String(Records("NAME").Value)
	kvo_auc=String(Records("KVO").Value)
	Records.MoveNext()
%>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="2" bgcolor="#FFFFFF" bordercolor="#CDE0DE"> 
            <p><a href="auction.asp?subj_id=<%=sb_id%>"> <%=name_subj%></a></p>
          </td>
          <td colspan="2" height="2" width="141" bordercolor="#CDE0DE"> 
            <p align="center"><%=kvo_auc%></p>
          </td>
        </tr>
        <%} Records.Close()%>
      </table>
    </td>
  </tr>
  <tr> 
    <td><font color="#74BADC">...................... ...................... ...................... 
      ...................... ...................... ...................... ...................... 
      . . . . .. .. . . .. . . .. .</font></td>
  </tr>
</table>
<p align="center"><img src="Images/arrow_rt.gif" width="20" height="60"> 
  <script language="JavaScript" src="http://vbn.tyumen.ru/cgi-bin/hints.cgi?vbn&bloka">
</script>
  <Script Language="JavaScript">
var code='';
</script>
  <script language="JavaScript" src="http://banners.isurgut.ru/GetBanner.asp?Member_id=296&Type_id=3"></script>
  <Script Language="JavaScript">
document.write(code);
</script>
</p>
<%}

if ((subj>0)&&(lot==0)) {%>
<table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#FFEFDF" bordercolor="#FFEFDF"> 
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
    <td>&nbsp; </td>
  </tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909">
  <tr bordercolor="#FFEFDF"> 
    <td width="150" bgcolor="#FFEFDF"> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan="2" bgcolor="#FFEFDF"> 
      <h1><b>:: <a href="auction.asp"><font face="Verdana, Arial, Helvetica, sans-serif">КАТАЛОГ</font></a> 
        </b>&gt;&gt; <%=name_subj%> &lt;&lt; <a href="add_lot.asp?subj_id=<%=subj%>">Добавить 
        лот</a></h1>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td>&nbsp; </td>
  </tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909" height="100">
  <tr valign="top"> 
    <td width="150" bordercolor="#FFEFDF" bgcolor="#FFEFDF" background="Images/line_back.gif"> 
      <h1 align="center"><b>Лоты</b></h1>
      <p>&nbsp;</p>
      <p><b>Ваши действия:</b></p>
      <p>&nbsp;</p>
      <p><a href="add_lot.asp?subj_id=<%=subj%>"><b>Добавить лот</b></a></p>
      <p><a href="reg_mem_auction.asp">Регистрация участника</a></p>
      <p><a href="auction_rules.html">Правила аукционов</a></p>
      <p><a href="area.asp">Личный кабинет</a></p>
      <p></p>
      <p><a href="reconnect.asp">Войти под другим именем</a> </p>
    </td>
    <td colspan="2" bgcolor="#FFFFFF" bordercolor="#FFFFFF" height="86"> 
      <table width="120" border="1" cellspacing="2" cellpadding="0" align="right" bordercolor="#330000">
        <tr> 
          <td bgcolor="#FFEFDF"> 
            <div align="CENTER"> 
              <h1><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Реклама</font></b></h1>
            </div>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <table border="0" cellspacing="2" cellpadding="0">
        <tr bgcolor="#FFEFDF"> 
          <td> 
            <p><b><font color="#000000">Название лота</font></b></p>
          </td>
          <td width="105"> 
            <p align="center"><b>Город</b></p>
          </td>
          <td valign="middle" align="center" width="110"> 
            <p><b>цена лота</b></p>
          </td>
          <td valign="middle" align="center" width="65"> 
            <p><b>тип</b></p>
          </td>
          <td valign="middle" align="center" width="45"> 
            <p><b>ставки</b></p>
          </td>
          <td valign="middle" align="center" width="85"> 
            <p><b>завершается</b></p>
          </td>
        </tr>
        <%
sql=select_goods
sql=sql.replace("%SUBJ_ID",subj)
sql=sql.replace("%SUBJ_ID",subj)
Records.Source=sql
Records.Open()
while (!Records.EOF)
{
	lot_id=String(Records("ID").Value)
	name_lot=String(Records("NAME").Value)
	tek_price=String(Records("PRICE").Value)
	valuta=String(Records("VALUTA").Value)
	kvo_memb=String(Records("KVO").Value)
	tip_nam=String(Records("type_name").Value)
	tip=String(Records("AUCTION_TYPE_ID").Value)
	city=String(Records("CITY").Value)
	end_date=Records("DATE_END").Value
   	if (tip==1) {nam_img="auc_regular"}
   	if (tip==2) {nam_img="auc_reserv"}
   	if (tip==4) {nam_img="auc_fix"}
   	if (tip==3) {nam_img="auc_inout"} 
	Records.MoveNext()
%>
        <tr bgcolor="#EBF3F2" valign="top"> 
          <td bgcolor="#EBF3F2"> 
            <p><font size="1">(лот № <%=lot_id%>)</font> <a href="auction.asp?subj_id=<%=subj%>&lot_id=<%=lot_id%>"><%=name_lot%></a></p>
          </td>
          <td width="105" bgcolor="#EBF3F2"> 
            <p><%=city%></p>
          </td>
          <td width="110" height="11" bgcolor="#EBF3F2"> 
            <p><font color="#3333FF"><b><%=tek_price%></b></font> <b><font color="#3E0909"><%=valuta%></font></b></p>
          </td>
          <td width="65" height="11"> <img src="<%="Images/"+nam_img+".gif"%>" width="65" height="18" align="absmiddle" alt="<%=tip_nam%>"> 
          </td>
          <td width="45" height="11" bgcolor="#EBF3F2" align="center"> 
            <p><%=kvo_memb%></p>
          </td>
          <td width="85" height="11" align="center"> 
            <p><%=end_date%></p>
          </td>
        </tr>
        <%} Records.Close()%>
      </table>
      <hr noshade size="1">
      <p><b><a href="auction_rules.html#1">Типы аукционов</a>:</b></p>
      <p><i><img src="Images/auc_regular.gif" width="65" height="18" align="absmiddle"> 
        "Обычный" - продажа по максимальной предложенной цене</i></p>
      <p><i><img src="Images/auc_reserv.gif" width="65" height="18" align="absmiddle"> 
        "С Зарезервированной Ценой" - продажа при условии достижения скрытой цены</i></p>
      <p><i><img src="Images/auc_fix.gif" width="65" height="18" align="absmiddle"> 
        &quot;С Фиксированной Ценой&quot; - продажа или покупка по первому предложению</i></p>
      <p><i><img src="Images/auc_inout.gif" width="65" height="18" align="absmiddle"> 
        "Наоборот" - покупка по наименьшей предложенной цене</i></p>
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
      <Script Language="JavaScript">
var code='';
</script>
      <script language="JavaScript" src="http://banners.isurgut.ru/GetBanner.asp?Member_id=296&Type_id=3"></script>
      <Script Language="JavaScript">
document.write(code);
</script>
    </td>
    <td height=60> 
      <script language="JavaScript" src="http://vbn.tyumen.ru/cgi-bin/hints.cgi?vbn&bloka">
</script>
    </td>
  </tr>
  </tbody> 
</table>
<%}
if ((subj>0)&&(lot>0)) {%>
<table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#FFEFDF" bordercolor="#FFEFDF"> 
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
    <td>&nbsp; </td>
  </tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909">
  <tr bordercolor="#FFEFDF"> 
    <td width="150" bgcolor="#FFEFDF"> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan="2" bgcolor="#FFEFDF"> 
      <h1><b>:: <a href="auction.asp"><font face="Verdana, Arial, Helvetica, sans-serif">КАТАЛОГ 
        </font></a></b>&gt;&gt; <a href="auction.asp?subj_id=<%=subj%>"><%=name_subj%></a> 
        &gt;&gt; <%=name_lot%> &lt;&lt; <a href="do_rate.asp?subj=<%=subj%>&lot=<%=lot%>">Предложите 
        свою цену</a></h1>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td>&nbsp; </td>
  </tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909">
  <tr valign="top"> 
    <td width="150" bordercolor="#FFEFDF" bgcolor="#FFEFDF" background="Images/line_back.gif"> 
      <h1 align="center"><b>Описание лота</b></h1>
      <p>&nbsp;</p>
      <p><b>Ваши действия:</b></p>
      <p>&nbsp;</p>
      <p><a href="do_rate.asp?subj=<%=subj%>&lot=<%=lot%>"><b>Сделать ставку </b></a></p>
      <p><a href="reg_mem_auction.asp">Регистрация участника</a></p>
      <p><a href="auction_rules.html">Правила аукционов</a></p>
      <p><a href="area.asp">Личный кабинет</a></p>
      <p><a href="reconnect.asp">Войти под другим именем</a> </p>
    </td>
    <td colspan="2" bgcolor="#FFFFFF" bordercolor="#FFFFFF"> 
      <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr> 
          <td bgcolor="#EBF3F2" colspan="2"> 
            <div align="center"> 
              <p align="left"><font size="4"><b><%=name_lot%></b></font> (лот 
                № <%=lot%>) <b><img src="Images/<%=nam_img%>.gif" width="65" height="18" align="absmiddle" alt="<%=tip_nam%>"> 
                </b></p>
            </div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFEFDF" colspan="2"> 
            <p><b>Текущая цена лота:</b><font color="#3333FF"> <b><font color="#FF0000"><%=tek_price%></font> 
              <%=valuta%> </b></font><%=kvo_stavok==0?" Ставок нет!":" Уже сделано "+kvo_stavok+" ставок."%> 
              <a href="do_rate.asp?subj=<%=subj%>&lot=<%=lot%>">Сделать ставку</a> 
              | <a href="Asklst.asp?subj_id=<%=subj%>&lot_id=<%=lot%>">Вопросы 
              по лоту</a> (<%=kask%>, новых : <b><font color="#FF0000"><%=newask%></font></b>)</p>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#EBF3F2" valign="top" colspan="2"> 
            <p><b>Владелец лота:</b> <b><font color="#000066"><%=host_lot%></font></b></p>
            <p><b>Город: <%=city%></b></p>
            <p><b>Описание лота:</b> </p>
            <table width="80%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td> 
                  <%
	if (filename!="") {
 		while (!ts.AtEndOfStream){
		str=ts.ReadLine()
		while(str.indexOf("<")>=0){str=str.replace("<","&lt;")}		
		Response.Write(str+"<br>")}
 		ts.Close()
	}
%>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFEFDF" width="170"> 
            <p><b>Начало торгов:<font color="#000066"> <%=start_date%></font></b></p>
          </td>
          <td bgcolor="#FFEFDF"> 
            <p><b>Окончание торгов: <font color="#FF0000"><%=end_date%></font></b></p>
          </td>
        </tr>
        <tr> 
          <td colspan="2" bgcolor="#EBF3F2"> 
            <p><img src="<%=image_name%>" align="absmiddle" alt="Иллюстрация к лоту"></p>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td bgcolor="#FFEFDF"> 
            <p><b>Управление лотом</b></p>
          </td>
        </tr>
        <tr> 
          <td valign="top"> 
            <p align="left"> <a href="edit_lot.asp?subj=<%=subj%>&lot=<%=lot%>"> 
              Изменить</a> | <a href="del_lot.asp?subj=<%=subj%>&lot=<%=lot%>">Удалить</a> 
              | <a href="prolong_lot.asp?subj=<%=subj%>&lot=<%=lot%>">Продлить</a> 
              | <a href="stop_lot.asp?subj=<%=subj%>&lot=<%=lot%>">Остановить</a></p>
          </td>
        </tr>
      </table>
      <hr noshade size="1">
      <p><b><a href="auction_rules.html#1">Типы аукционов</a>:</b></p>
      <p><i><img src="Images/auc_regular.gif" width="65" height="18" align="absmiddle"> 
        "Обычный" - продажа по максимальной предложенной цене</i></p>
      <p><i><img src="Images/auc_reserv.gif" width="65" height="18" align="absmiddle"> 
        "С Зарезервированной Ценой" - продажа при условии достижения скрытой цены</i></p>
      <p><i><img src="Images/auc_fix.gif" width="65" height="18" align="absmiddle"> 
        &quot;С Фиксированной Ценой&quot; - продажа или покупка по первому предложению</i></p>
      <p><i><img src="Images/auc_inout.gif" width="65" height="18" align="absmiddle"> 
        "Наоборот" - покупка по наименьшей предложенной цене</i></p>
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
      <Script Language="JavaScript">
var code='';
</script>
      <script language="JavaScript" src="http://banners.isurgut.ru/GetBanner.asp?Member_id=296&Type_id=3"></script>
      <Script Language="JavaScript">
document.write(code);
</script>
    </td>
    <td height=60 bgcolor="#3E0909"> 
      <script language="JavaScript" src="http://vbn.tyumen.ru/cgi-bin/hints.cgi?vbn&bloka">
</script>
    </td>
  </tr>
  </tbody> 
</table>
<%}%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909">
  <tr bgcolor="#FFEFDF" bordercolor="#FFEFDF"> 
    <td height="24" bgcolor="#CDE0DE" bordercolor="#CDE0DE"> 
      <h1 align="center"><font size="1">Программирование и дизайн <a href="http://www.rusintel.ru/" target="_blank">ЗАО 
        Русинтел</a> &copy; 2001-2002</font></h1>
    </td>
  </tr>
</table>
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
width="1" height="1" ></a></noscript> <!-- /HotLog --> 
</body>
</html>
