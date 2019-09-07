<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\path.inc" -->


<%
var lot_id=parseInt(Session("lot_id"))
var subj=parseInt(Request("subj_id"))

if(String(Session("id_mem"))=="undefined"){
 Session("backurl")="add_lot_img.asp?subj_id="+subj
 Response.Redirect("login.asp")
}

var ErrorMsg=""
var subj_name=""
var lot_name=""
var image_name=""


if (isNaN(lot_id)) {Response.Redirect("auction.asp")}

if (isNaN(subj)) {
	if (Request.ServerVariables("REQUEST_METHOD")=="POST") {
	   updown = Server.CreateObject("ANUPLOAD.OBJ")
		subj=updown.Form("subj_id")
		if (!isNaN(parseInt(subj))) {
		try {
			updown.Delete(LotsFilePath+lot_id+".gif")
			updown.Delete(LotsFilePath+lot_id+".jpg")
			updown.SavePath = LotsFilePath
			size=parseInt(updown.GetSize("file"))
			ext=updown.GetExtension("file").toUpperCase()
			if(ext!="JPG" && ext!="GIF"){throw "Принимаются только JPG или GIF файлы."}
			if (size>20480){throw "Не более 20kB."}
			if (size==0){throw "Нет файла."}
			updown.SaveAs("file",LotsFilePath+lot_id+"."+ext)
			Response.Redirect("auction.asp?subj_id="+subj+"&lot_id="+lot_id)
			}
    	catch(e){ErrorMsg+=String(e.message)=="undefined"?e:e.message}
		
		} else {Response.Redirect("auction.asp")}
	}
	else {Response.Redirect("auction.asp")}
}

var fs= new ActiveXObject("Scripting.FileSystemObject")

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


if(Session("state_mem")!=0){
 Session("backurl")="add_lot_img.asp?subj_id="+subj
 Response.Redirect("activete_account.asp")
}

if (isNaN(lot_id)) {Response.Redirect("auction.asp?subj_id="+subj)}

Records.Source="Select * from GOODS where ID="+lot_id
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("auction.asp?subj_id="+subj)
}
lot_name=String(Records("NAME").Value)
Records.Close()
LotsFilePath
image_name=GlobLotsFilePath+lot_id+".jpg"
if(!fs.FileExists(LotsFilePath+lot_id+".jpg")){ image_name="" }
if(image_name==""){
	image_name=GlobLotsFilePath+lot_id+".gif"
	if(!fs.FileExists(LotsFilePath+lot_id+".gif")){ image_name="" }
}


%>
<HTML>
<HEAD>
<style type="text/css">
<!--
p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 10pt; font-weight: 400; color: #333333; margin:  3px}
a:hover {color:#FF3333}
a:visited {  color: #333333}
h1 {  font: 10pt/normal Verdana, Arial, Helvetica, sans-serif; color: #990000; text-decoration: blink; border-width: 1px auto medium medium; border-color: #CCCCFF black black; background: repeat-y; margin: 3px 3px}
a:link {  color: #990000}
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<title>Аукцион Нужных Вещей</title>
<link rel="stylesheet" href="/auction.css" type="text/css">
</HEAD>
<BODY BGCOLOR="#ffffff" background="/Images/block_back.gif">
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
      size=1 color="#CC9900"><b>Аукцион Полезных Вещей</b></font></font></bl></font></div>
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
<table border=1 bordercolor=#3E0909 cellpadding=0 cellspacing=0 width="100%">
  <tbody> 
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td width=150> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan=2> 
      <p>:: <a href="auction.asp">Каталог</a> | <a href="reg_mem_auction.asp">Регистрация</a> 
        | <a href="area.asp">Личный кабинет</a> | <a href="reconnect.asp">Войти 
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
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td width="150" height="24"> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan="2" height="24"> 
      <h1> <b>:: Изменение параметров лота::</b></h1>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#3E0909" width="60">&nbsp;</td>
    <td>&nbsp; </td>
  </tr>
</table>
<p><br>
</p>
<FORM action="add_lot_img.asp" name="form" method="post" enctype="multipart/form-data" OnSubmit="return Check()">
  <input type="hidden" name="subj_id" value="<%=subj%>">
  <center class="doctitle">
    Изменение параметров лота. 
  </center> 
   <%if(ErrorMsg!=""){%> 
  <h3 align="center"> <font color="#FF3333">Ошибка!</font><font color="#FF3333"><br>
    </font><font color="#1e4a7b"><%=ErrorMsg%></font><br>
  </h3>
<%}%>
  <br>
  <table width="100%" border="1" cellspacing="0" bordercolorlight="#CCCCCC" bordercolordark="#666666">
    <tr> 
      <td height="86" bgcolor="#FFFFFF"> 
        <table width="100%" border="0" cellspacing="3">
          <tr> 
            <td colspan="2" class="blueband">
              <p>Информация о лоте</p>
              </td>
          </tr>
          <tr> 
            <td class="fieldtitle" width="20%" height="2">Наименование лота:</td>
            <td class="fielddata" height="2">
              <p><b><font color="#0066CC"><%=lot_name%></font></b></p>
            </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="3">
          <tr> 
            <td class="blueband" colspan="2" bgcolor="#FFFFFF"> 
              <p>Фотография (риснок) лота</p>
            </td>
          </tr>
          <tr class="fieldtitle_L">
            <td colspan="2" class="fieldtitle_L" height="22" bgcolor="#FFFFFF" >
<center><img src="<%=image_name==""?"lots/noimg.jpg":image_name%>"> </center>
            </td>
          </tr>
          <tr class="fieldtitle_L"> 
            <td colspan="2" class="fieldtitle_L" height="2" bgcolor="#FFFFFF" > 
              <div align="left">Файл с фотографией: 
                <input type="file" name="file" size="50">
              </div>
            </td>
          </tr>
        </table>
        <p align="center">
          <input type="submit" name="Submit" value="Установить" >
        </p>
        </td>
    </tr>
  </table>
  </FORM>

<div align="right">
  <p align="center"> </p>
</div>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#330000">
  <tr bordercolor="#FFCC99"> 
    <td width="150" bgcolor="#FFCC99"> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan="2" bgcolor="#FFCC99"> 
      <h1><b><a href="auction.asp">КАТАЛОГ</a> &gt;&gt;<font color="#6699FF"><a href="auction.asp?subj_id=<%=subj%>"><%=subj_name%></a></font></b></h1>
    </td>
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
<p>&nbsp;</p>
</BODY>
</HTML>
