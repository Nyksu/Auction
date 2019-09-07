<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\sql.inc" -->

<%
var ErrorMsg=""
var sql=""
var subj=parseInt(Request("subj_id"))
var subj_name=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var txt=""
var stepz=2
var Name=""
var valut=""


if (isNaN(subj)) {Response.Redirect("auction.asp")}

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
 Session("backurl")="add_lot.asp?subj_id="+subj
 Response.Redirect("login.asp")
}

if(Session("state_mem")!=0){
 Session("backurl")="add_lot.asp?subj_id="+subj
 Response.Redirect("activete_account.asp")
}

isFirst=String(Request.Form("add_auction"))=="undefined"
ShowForm=true

if(!isFirst){
		Name=TextFormData(Request.Form("lot_name"),"")
		valut=TextFormData(Request.Form("curanc"),"")
		txt=TextFormData(Request.Form("descr"),"")
		st_price=TextFormData(Request.Form("start_price"),"")
		typeid=IntFormData(Request.Form("type_auc"))
		d_end=TextFormData(Request.Form("exp_date"),"")
		stepz=IntFormData(Request.Form("steps"))
		
		if (typeid==2){price=FloatFormData(Request.Form("min_price"))}
		if ((typeid==2)&&(price<st_price)){ErrorMsg+="'Минимальная цена' должна быть больше Стартовой стоимости лота.<br>"}
		if (!/(\d{1,2}).(\d{1,2}).(\d{4})$/.test(d_end)){ErrorMsg+="'Дата окончания' некорректная.<br>"}
		if(Name.length<2){ErrorMsg+="'Наименование лота' должно быть больше 1-го символа.<br>"}
		if((txt.length<8)||(txt.length>3000)){ErrorMsg+="Описание лота не может быть слишком коротким и слишком длинным.<br>"}
		if (st_price==""){ErrorMsg+="'Стартовая стоимость лота' не заполнено.<br>"}
		st_price=parseFloat(st_price,10)
		if (isNaN(st_price)){ErrorMsg+="'Стартовая стоимость лота' неверного формата.<br>"}
		if (typeid!=2){price=st_price}
		price=parseFloat(price,10)
		if (isNaN(price)){ErrorMsg+="'Минимальная цена' неверного формата.<br>"}
		if((valut=="")||(valut==" ")){ErrorMsg+="'Валюта' некорректная.<br>"}
		if (stepz>50 || stepz<1) {ErrorMsg+="'Шаг ставок' должен быть от 1 до 50 % !<br>"}


		if (ErrorMsg==""){
			ID_LOT=NextID("UNIVERSAL")
			filename=LotsFilePath+ID_LOT+LotsExtention
			while (Name.indexOf("'")>=0) {Name=Name.replace("'","\"")}
			while (valut.indexOf("'")>=0) {valut=valut.replace("'","\"")}
			sql=insert_lot
			sql=sql.replace("%ID",ID_LOT)
			sql=sql.replace("%NAME",Name)
			sql=sql.replace("%MID",Session("id_mem"))
			sql=sql.replace("%SUBJID",subj)
			sql=sql.replace("%TYPEID",typeid)
			sql=sql.replace("%STARTPRICE",st_price)
			sql=sql.replace("%PRICE",st_price)
			sql=sql.replace("%ZPRICE",price)
			sql=sql.replace("%DATEND",d_end)
			sql=sql.replace("%VALUTA",valut)
			sql=sql.replace("%DD",stepz)
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
				Session("lot_id")=ID_LOT
			 }
			 catch(e){
				ErrorMsg+=filename
				Connect.RollbackTrans()
				}
			}
		}
while (Name.indexOf("<")>=0) {Name=Name.replace("<","&lt;")}
while (Name.indexOf("\"")>=0) {Name=Name.replace("\"","&quot;")}
while (valut.indexOf("<")>=0) {valut=valut.replace("<","")}
while (valut.indexOf("\"")>=0) {valut=valut.replace("\"","&quot;")}
}


%>

<html>
<head>
<title>Аукцион Полезных Вещей</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<style type="text/css">
<!--
p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 10pt; font-weight: 400; color: #000000; margin:  3px}
a:hover {color:#FF3333; font: 9pt/9pt Arial, Helvetica, sans-serif}
a:visited {  font-family: Arial, Helvetica, sans-serif; color: #990000; font-size: 9pt; line-height: 9pt}
h1 {  font: 10pt/normal Verdana, Arial, Helvetica, sans-serif; color: #3E0909; text-decoration: blink; border-width: 1px auto medium medium; border-color: #CCCCFF black black; background: repeat-y; margin: 3px 3px}
a:link {  font: 9pt/9pt Arial, Helvetica, sans-serif; color: #000000}
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
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#FFCC99" width="150" valign="top" bordercolor="#330000">&nbsp;</td>
    <td bgcolor="#FFFFFF">
	<%if(ErrorMsg!=""){%>
  <center>
    <h2> 
      <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
        <%=ErrorMsg%></p>
    </h2>
  </center>
  <%}%>
  
<form name="form1" method="post" action="add_lot.asp">
  <input type="hidden" name="subj_id" value="<%=subj%>">
    
    <%if (ShowForm) {%>
          <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909">
            <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
              <td width="150"> 
                <h1>&nbsp;</h1>
        </td>
              <td colspan="2">
              <h1>Тема аукционов: <a href="auction.asp?subj_id=<%=subj%>"><%=subj_name%></a> 
              </h1>
      </td>
      </tr>
    </table>
    
  <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
    <tr bordercolor="#CCCCCC"> 
      <td width="150" bgcolor="#CCCCCC"> 
        <h1>&nbsp;</h1>
        </td>
        
      <td colspan="2" bgcolor="#CCCCCC"> 
        <h1> Форма регистрации лота.</h1>
        </td>
      </tr>
    </table>
    
  <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909">
    <tr valign="top"> 
        <td bordercolor="#FFFFFF" height="378" width="863"> 
          
              <table width="100%" border="0" bgcolor="#FFFFFF">
                <tr> 
                  <td> 
                    <p><b>Перед заполнением данной формы Обязательно ознакомтесь 
                      с </b><a href="auction_rules.html">Правилами Сибирского 
                      Аукциона</a>!</p>
                    <p><b>Заполняя форму Вы соглашаетесь со всеми правилами Аукционов!</b></p>
                    <table width="100%" border="0" cellspacing="4" cellpadding="0" bgcolor="#FFFFFF">
                      <tr valign="middle"> 
                        <td width="175" align="right" valign="middle" bgcolor="#E0DEFE" height="24"> 
                          <p><b>Выберите тип аукциона<font size="3">*</font></b></p>
                        </td>
                        <td bgcolor="#E0DEFE" height="24" width="580"> 
                          <p> 
                            <select name="type_auc">
                              <% 		Records.Source="select * from auction_type"
		Records.Open()
		while(!Records.EOF){%>
                              <option value="<%=Records("ID").Value%>"
			<%=isFirst&&Records("ID").Value==1?"selected":""%>
			<%=!isFirst&&Records("ID").Value==Request.Form("type_auc")?"selected":""%>> 
                              <%=Records("NAME").Value%></option>
                              <%Records.MoveNext()
		}
                Records.Close()%>
                            </select>
                            <a href="auction_rules.html" target="_blank">подробнее 
                            о типах аукционов</a></p>
                        </td>
                      </tr>
                      <tr valign="middle"> 
                        <td width="175" align="right" valign="middle" bgcolor="#EEEDFE"> 
                          <p><b>Наименование лота <font size="3">*</font> </b></p>
                        </td>
                        <td bgcolor="#EEEDFE" width="580"> 
                          <input type="text" name="lot_name" size="40" maxlength="80" value="<%=isFirst?"":Request.Form("lot_name")%>">
                        </td>
                      </tr>
                      <tr valign="top"> 
                        <td width="175" align="right" valign="middle" bgcolor="#E0DEFE"> 
                          <p align="right"><b>Описание Лота <font size="3">*</font> 
                            </b></p>
                          <p align="left"><br>
                            Подробное описание лота повысит вероятность успешного 
                            завершения вашего аукциона</p>
                        </td>
                        <td bgcolor="#E0DEFE" width="580"> 
                          <textarea name="descr" cols="60" rows="6"><%=txt%></textarea>
                        </td>
                      </tr>
                      <tr valign="middle"> 
                        <td width="175" align="right" valign="middle" bgcolor="#EEEDFE"> 
                          <p><b>Наименование валюты<font size="3">*</font></b></p>
                        </td>
                        <td bgcolor="#EEEDFE" width="580"> 
                          <input type="text" name="curanc" size="10" maxlength="10" value="<%=isFirst?"":Request.Form("curanc")%>">
                          (USD, $, руб. и т.д.) <font color="#0000FF">Только наименование 
                          валюты!!</font></td>
                      </tr>
                      <tr> 
                        <td width="175" height="7" align="right" valign="middle" bgcolor="#E0DEFE"> 
                          <p><b>Стартовая стоимость лота<font size="3">*</font></b></p>
                        </td>
                        <td height="7" bgcolor="#E0DEFE" width="580"> 
                          <input type="text" name="start_price" size="20" maxlength="20" value="<%=isFirst?"":Request.Form("start_price")%>">
                          (без пробелов)</td>
                      </tr>
                      <tr> 
                        <td width="175" height="6" valign="middle" bgcolor="#EEEDFE"> 
                          <p align="right"><b>Минимальная цена </b></p>
                        </td>
                        <td height="6" valign="top" bgcolor="#EEEDFE" width="580"> 
                          <input type="text" name="min_price" size="20" maxlength="20" value="<%=isFirst?"":Request.Form("min_price")%>">
                          (только для аукциона с зарезервированной ценой <b>*</b>) 
                        </td>
                      </tr>
                      <tr> 
                        <td width="175" height="5" valign="middle" bgcolor="#E0DEFE"> 
                          <p align="right"><b>Дата окончания аукциона <font size="3">*</font></b></p>
                        </td>
                        <td height="5" valign="top" bgcolor="#E0DEFE" width="580"> 
                          <input type="text" name="exp_date" size="20" maxlength="20" value="<%=isFirst?"":Request.Form("exp_date")%>">
                          <p>(dd.mm.yyyy) где dd - день, mm - месяц, yyyy - год. 
                          </p>
                          <p>Разделительный знак - точка (.) обязателен!</p>
                        </td>
                      </tr>
                      <tr> 
                        <td width="175" height="22" valign="middle" bgcolor="#EEEDFE"> 
                          <div align="right"> 
                            <p><b>Шаг ставок</b><font size="3"> *</font></p>
                          </div>
                        </td>
                        <td height="22" bgcolor="#EEEDFE" valign="top" width="580"> 
                          <p> 
                            <input type="text" name="steps" maxlength="2" size="5" value="<%=isFirst?"2":Request.Form("steps")%>">
                            <font color="#0000FF"><b>%</b></font><b></b><font size="4"> 
                            </font>(определяет минимальное приращение ставки в 
                            процентах от стартовой цены и может быть в интервале 
                            от 1 до 50 %)</p>
                        </td>
                      </tr>
                      <tr> 
                        <td width="175" height="22" valign="top">&nbsp;</td>
                        <td height="22" width="580"> 
                          <p><b><font size="4">*</font> - Обязательные поля</b></p>
                        </td>
                      </tr>
                      <tr> 
                        <td width="175" height="22" valign="top"> 
                          <p>&nbsp;</p>
                        </td>
                        <td height="22" width="580"> 
                          <p> 
                            <input type="submit" name="add_auction" value="Регистрация">
                            <input type="reset" name="Submit2" value="Очистить">
                          </p>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
        </td>
      </tr>
    </table>
    
  <hr size="1">
        <p align="center">&nbsp;</p>
    <p align="center"> 
      <%}else {%>
    </p>
    <p></p>
    <p></p>
  </form>
  <div align="center"> 
    <p>&nbsp;</p>
    <p><font size="3" color="#0000FF"><b>Лот успешно зарегистрирован. </b></font></p>
  </div>
      <p align="center">&nbsp;</p>
      <p align="center"><b>Вы так же можете разместить иллюстрацию или фотографию 
        к описанию лота: <a href="add_lot_img.asp?subj_id=<%=subj%>">ЗДЕСЬ</a></b></p>
  
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
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#3E0909">
        <tr bordercolor="#FFCC99"> 
          <td width="150" bgcolor="#FFCC99"> 
            <h1>&nbsp;</h1>
    </td>
          <td colspan="2" bgcolor="#FFCC99"> 
            
      <h1><a href="auction.asp">КАТАЛОГ </a> &gt;&gt;<a href="auction.asp?subj_id=<%=subj%>"> 
        <%=subj_name%></a></h1>
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
