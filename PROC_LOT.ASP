<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\email.inc" -->

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
var isFirst=true
var start_price=0
var test_price=0
var kvo_stavok=0
var ShowForm=true
var txt=""
var isSending=false
var membername=""
var mememail=""
var memphone=""
var memnik=""
var z_price=0
var fs= new ActiveXObject("Scripting.FileSystemObject")
var eml=Server.CreateObject("JMail.Message")
var filename=LotsFilePath+lot+LotsExtention

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
 Session("backurl")="proc_lot.asp?subj="+subj+"&lot="+lot
 Response.Redirect("login.asp")
}

if(Session("state_mem")!=0){
 Session("backurl")="proc_lot.asp?subj="+subj+"&lot="+lot
 Response.Redirect("activete_account.asp")
}

if (lot>0){
   sql=select1_allfor_lots
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
   z_price=Records("Z_PRICE").Value

   Records.Close()
    
   if (host_lot!=Session("id_mem")){Response.Redirect("auction.asp?subj_id="+subj+"&lot_id="+lot)}
	
   if (atip==1) {nam_img="auc_regular"}
   if (atip==2) {nam_img="auc_reserv"}
   if (atip==4) {nam_img="auc_fix"}
   if (atip==3) {nam_img="auc_inout"}
   
   sql="Select Count(*) as kvo from trade_history where goods_id="+lot
	Records.Source=sql
	Records.Open()
	kvo_stavok=Records("KVO").Value
	Records.Close()
   
   if (kvo_stavok==0) {atensiontxt="Удаление повлечет за собой невозможность восстановления данных по этому лоту!!"}
   else {atensiontxt="По лоту размещены ставки!! (зарегистрировано ставок:  "+kvo_stavok+"). После удаления Вы не сможите использовать результаты аукциона!!"}
   
   if ((atip==2)&&(z_price<tek_price)){
      Session("backurl")="lots_end_list.asp"
   	  Response.Redirect("amsg.asp?msg=1")
   }
   
   sql=winner_lot
   sql=sql.replace("%LOT",lot)
   Records.Source=sql
   Records.Open()
    if (Records.EOF){
	  Records.Close()
	  Response.Redirect("lots_end_list.asp")
   }
   membername=Records("NAME").Value
   mememail=Records("E_MAIL").Value
   memphone=Records("PHONE").Value
   memnik=Records("NIK_NAME").Value
   Records.Close()  
}

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst){
			txt=TextFormData(Request.Form("msg"),"")
			if (txt.length>2000){ErrorMsg+="Сообщение превышает допустимый размер.<br>"}
			if ((TextFormData(Request.Form("is_msg"),"")!="1")){txt=""}

			if (ErrorMsg==""){
			  var  pop3=Server.CreateObject("JMail.POP3")
			  pop3.Connect(fromaddres,pswsmtp,servsmtp)
			  pop3.Disconnect()
			  eml.Logging=true
			  eml.From=fromaddres
			  eml.AddRecipient(mememail)
			  eml.Subject="Уведомление. Сибирский аукцион. Ваша ставка выйграла аукцион!"
			  eml.Charset=characterset
			  eml.ContentTransferEncoding = "base64"
			  eml.AppendText(" Здравствуйте! \n\n")
			  eml.AppendText(" Сибирский аукцион WWW.AUCTION.72RUS.RU искренне поздравляет Вас! \n")
			  eml.AppendText(" Ваша ставка выйграла в аукционе ( [["+name_lot+"]] лот № "+lot+")! \n")
			  eml.AppendText(" Это письмо активировано хозяином лота. В ближайшее время он свяжется с Вами. \n")
			  eml.AppendText(" Напоминаем, Вами заявлена ставка по этому лоту в размере: "+tek_price+" "+valuta+"\n")
			  eml.AppendText(" \n\n")
			  eml.AppendText("Сообщение от владельца лота: \n\n")
			  eml.AppendText(txt)
			  isSending=eml.Send(servsmtp)
	 		  if (!isSending) {ErrorMsg+="Почта не отправлена. Ошибка отправления.<br>"}
			  
			  if (ErrorMsg==""){
				sql=set_lot_proc
				sql=sql.replace("%LOT",lot)
				Connect.BeginTrans()
				try{
				  Connect.Execute(sql)
				}
				  catch(e){
				  Connect.RollbackTrans()
				  ErrorMsg+=ListAdoErrors()
				}
				if (ErrorMsg==""){
				  Connect.CommitTrans()
				  ShowForm=false
				}
			  }
			}
}

%>

<html>
<head>
<title>Сибирский Аукцион полезных вещей. Обработка лота.</title>
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
<link rel="stylesheet" href="/auction.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="/Images/block_back.gif">
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
      <p>:: <a href="reg_mem_auction.asp">Регистрация</a> | <a href="auction_rules.html">Правила 
        аукционов</a> | <a href="area.asp">Личный кабинет</a> | <a href="reconnect.asp">Войти 
        под другим именем</a> :: </p>
    </td>
  </tr>
  </tbody> 
</table>
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
<form name="form" method="post" action="proc_lot.asp">
  <p align="center"> 
    <input type="hidden" name="subj" value="<%=subj%>">
    <input type="hidden" name="lot" value="<%=lot%>">
  </p>
  <%if (ShowForm) {%>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#330000">
    <tr> 
      <td width="150" bordercolor="#FFCC99" bgcolor="#FFCC99"> 
        <h1>&nbsp;</h1>
      </td>
      <td colspan="2" bordercolor="#FFCC99" bgcolor="#FFCC99"> 
        <p><b><a href="auction.asp"><font face="Verdana, Arial, Helvetica, sans-serif">КАТАЛОГ 
          </font></a>&gt;&gt; <a href="auction.asp?subj_id=<%=subj%>"><%=subj_name%></a> &gt;&gt; <%=name_lot%> </b></p>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#3E0909" width="60">&nbsp;</td>
      <td>&nbsp; </td>
    </tr>
  </table>
  <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#330000">
    <tr valign="top"> 
      <td width="150" bordercolor="#FFCC99" bgcolor="#FFCC99" background="Images/line_back.gif"> 
        <h1 align="center"><b>Описание&gt;&gt;</b></h1>
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
            <td rowspan="6" width="120">&nbsp; </td>
          </tr>
          <tr> 
            <td bgcolor="#BDCECE" colspan="2" height="24"> 
              <p><b>Текущая цена лота:</b><font color="#3333FF"> <b><font color="#FF0000"><%=tek_price%></font> <%=valuta%> </b></font><%=kvo_stavok==0?" Ставок нет!":" Уже сделано "+kvo_stavok+" ставок."%></p>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#EBF3F2" height="4" valign="top" colspan="2"> 
              <p><b><font size="3" color="#FF0000">Внимание!! <%=atensiontxt%></font></b></p>
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
            <td bgcolor="#EBF3F2"> 
              <p align="right"><b>Ваше согласие:<font size="3" color="#FF0000">&nbsp;</font></b></p>
            </td>
            <td bgcolor="#EBF3F2"> 
              <p><font color="#3333FF"> 
                <input type="checkbox" name="is_msg" value="1">
                <b><font size="2"> Включить в извещение Ваше сообщение</font></b></font></p>
            </td>
          </tr>
          <tr> 
            <td> 
              <p align="right">&nbsp;</p>
            </td>
            <td> 
              <p align="left"> 
                <textarea name="msg" cols="80" rows="8"></textarea>
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <p align="center"><font color="#CC0000"><b>Эта форма извещает победившего участника, 
    а так же переводит лот в состояние &quot;ОБРАБОТАННЫЙ&quot; списывая его в 
    АРХИВ</b></font></p>
  <p align="center"><b><font color="#CC0000">После отправки этой формы вы получите 
    полную информацию как связаться с выйгравшим участником.</font></b></p>
  <p align="center"> 
    <input type="submit" name="Submit" value="Submit">
  </p>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#3E0909" width="60">&nbsp;</td>
      <td>&nbsp; </td>
    </tr>
  </table>
  <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909">
    <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
      <td height="24"> 
        <h1 align="center"><font size="1">Программирование и дизайн <a href="http://www.rusintel.ru/" target="_blank">ЗАО 
          Русинтел</a> &copy; 2001-2002</font></h1>
      </td>
    </tr>
  </table>
  <p align="center"> 
    <%}else {%>
  </p>
</form>
<div align="right"> 
  <div align="center"> 
    <p><font size="3" color="#0000FF"><b>Лот обработан и перемещен в АРХИВ. </b></font></p>
  </div>
  <p align="center">&nbsp;</p>
  <p align="center"><b><font size="3">Контактные данные о выйгравшем участнике:</font></b></p>
  <div align="center"> 
    <table width="50%" border="1" bordercolor="#FFFFFF">
      <tr bordercolor="#3399CC"> 
        <td height="21" width="40%"> 
          <p><b>Псевдоним:</b></p>
        </td>
        <td height="21"> 
          <p><b><font color="#333399"><font size="2"><%=memnik%></font></font></b></p>
        </td>
      </tr>
      <tr bordercolor="#3399CC"> 
        <td width="40%"> 
          <p><b>Ф.И.О. (название организации):</b></p>
        </td>
        <td> 
          <p><b><font color="#333399"><font size="2"><font size="2"><b><%=membername%></b></font></font></font></b></p>
        </td>
      </tr>
      <tr bordercolor="#3399CC"> 
        <td width="40%"> 
          <p><b>Телефон:</b></p>
        </td>
        <td> 
          <p><b><font color="#333399"><font size="2"><font size="2"><b><%=memphone%></b></font></font></font></b></p>
        </td>
      </tr>
      <tr bordercolor="#3399CC"> 
        <td width="40%"> 
          <p><b>E-mail:</b></p>
        </td>
        <td> 
          <p><b><font color="#333399"><font size="2"><font size="2"><b><%=mememail%></b></font></font></font></b></p>
        </td>
      </tr>
    </table>
  </div>
  <p align="center">&nbsp;</p>
</div>
<%}%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC99">
  <tr> 
    <td width="150" bordercolor="#FFCC99" bgcolor="#FFCC99"> 
      <h1><a href="area.asp">Кабинет</a></h1>
    </td>
    <td colspan="2" bordercolor="#FFCC99" bgcolor="#FFCC99"> 
      <p><b><a href="auction.asp"><b><font face="Verdana, Arial, Helvetica, sans-serif">Весь 
        КАТАЛОГ </font></b></a> &gt;&gt;<font color="#6699FF"><a href="auction.asp?subj_id=<%=subj%>"><%=subj_name%></a></font></b></p>
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
