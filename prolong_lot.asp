<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var ErrorMsg=""
var sql=""
var subj=parseInt(Request("subj"))
var lot=parseInt(Request("lot"))
var subj_name=""
var res=""
var name_lot=""
var tek_price=0
var valuta=""
var kvo_stavok=0
var host_lot=""
var end_date=""
var start_date=""
var nam_img=""
var tip_nam=""
var atip=0
var isFirst=true
var ShowForm=true
var sumdat=Server.CreateObject("datesum.DateSummer")

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
 Session("backurl")="prolong_lot.asp?subj="+subj+"&lot="+lot
 Response.Redirect("login.asp")
}

if(Session("state_mem")!=0){
 Session("backurl")="prolong_lot.asp?subj="+subj+"&lot="+lot
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
   atip=Records("AUCTION_TYPE_ID").Value
   start_price=Records("START_PRICE").Value
   tip_nam=Records("type_name").Value

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
   
  
}

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst){
	if (TextFormData(Request.Form("is_del"),"")=="1"){
		d_end=TextFormData(Request.Form("exp_date"),"")
		if (!/(\d{1,2}).(\d{1,2}).(\d{4})$/.test(d_end)){ErrorMsg+="'���� ���������' ������������.<br>"}
		if (ErrorMsg=="") {
			res=sumdat.DateComparing(d_end,end_date)
			if (res==null) {ErrorMsg+="'���� ���������' ������������.<br>"}
			else {
				if (res<=0) {ErrorMsg+="���� ��������� ������ ���� ������ ���� ��������� ��������.<br>"}
			}
		}
		if (ErrorMsg=="") {
			sql=set_lot_prolong
			sql=sql.replace("%LOT",lot)
			sql=sql.replace("%DE",d_end)
			Connect.BeginTrans()
			try{
				Connect.Execute(sql)
				Connect.CommitTrans()
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+="������ �������!!.<br>"
				ErrorMsg+=ListAdoErrors()
			}
			if (ErrorMsg==""){ShowForm=false}
		}
	}
	else {Response.Redirect("auction.asp?subj_id="+subj+"&lot_id="+lot)}
}

%>

<html>
<head>
<title>��������� ������� �������� �����. ��������� ����.</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<style type="text/css">
<!--
p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 10pt; font-weight: 400; color: #333333; margin:  3px}
a:hover {color:#FF3333; font-family: Arial, Helvetica, sans-serif}
a:visited {  font-family: Arial, Helvetica, sans-serif; color: #990000}
h1 {  font: 10pt/normal Verdana, Arial, Helvetica, sans-serif; color: #990000; text-decoration: blink; border-width: 1px auto medium medium; border-color: #CCCCFF black black; background: repeat-y; margin: 3px 3px}
a:link {  color: #333333; font-family: Arial, Helvetica, sans-serif}
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
      size=6><bl><font color="#FFCC99" face="Times New Roman, Times, serif">��������� 
        �������</font><font color="#FFFFFF"><br>
        <font face="Arial, Helvetica, sans-serif" 
      size=6><font color=#FFFFFF face="Arial, Helvetica, sans-serif" 
      size=6><font color=#FFFFFF face="Arial, Helvetica, sans-serif" 
      size=6><b><font size="1" color="#CC9900">������� �������� ����� / </font><font color=#FFFFFF face="Verdana, Arial, Helvetica, sans-serif" 
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
var dayarray=new Array("�����������","�����������","�������","�����","�������","�������","�������")
var montharray=new Array("������","�������","����","������","���","����","����","������","��������","�������","������","�������")
document.write("<small><font color='#CC9900' face='Arial'><b>"+dayarray[day]+", "+montharray[month]+" "+daym+", "+year+"</b></font></small>")

</script>
        </font></b></font></font></font></font></bl></font></div>
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
      <p>:: <a href="auction.asp">�������</a> | <a href="reg_mem_auction.asp">�����������</a> 
        | <a href="area.asp">������ �������</a> | <a href="reconnect.asp">����� 
        ��� ������ ������</a> :: </p>
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
      <h1> <b>:: ��������� ���� ::</b></h1>
    </td>
  </tr>
</table>
<div align="right"><p align="center"> 
    <%if(ErrorMsg!=""){%>
  </p>
  <center>
    <h2> 
      <p> <font color="#FF3300" size="2"><b>������!</b></font> <br>
        <%=ErrorMsg%></p>
    </h2>
  </center>
  <%}%>
</div>
<form name="form1" method="post" action="prolong_lot.asp">
  <p align="center"> 
    <input type="hidden" name="subj" value="<%=subj%>">
    <input type="hidden" name="lot" value="<%=lot%>">
  </p>
    <%if (ShowForm) {%>
  <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#333333">
    <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
      <td width="150" bgcolor="#FFCC99"> 
        <h1>&nbsp;</h1>
      </td>
      <td colspan="2"> 
        <p><b><a href="auction.asp"><font face="Verdana, Arial, Helvetica, sans-serif">���� 
          �������</font></a> &gt;&gt; <a href="auction.asp?subj_id=<%=subj%>"><%=subj_name%></a> 
          &gt;&gt; <%=name_lot%> </b></p>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#3E0909" width="60">&nbsp;</td>
      <td>&nbsp; </td>
    </tr>
  </table>
  <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#333333">
    <tr valign="top"> 
      <td width="150" bordercolor="#FFCC99" height="148" bgcolor="#FFCC99" background="Images/line_back.gif"> 
        <h1 align="center"><b>��������</b></h1>
        <p>&nbsp;</p>
        </td>
      <td colspan="2" bgcolor="#FFFFFF" bordercolor="#FFFFFF" height="148"> 
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
          <tr> 
            <td bgcolor="#EBF3F2" colspan="2"> 
              <div align="center"> 
                <p align="left"><font size="1"><b><img src="<%="Images/"+nam_img+".gif"%>" height="16" align="absmiddle" alt="<%=tip_nam%>"></b> 
                  (��� � <%=lot%>)</font> <a href="auction.asp?subj_id=<%=subj%>&lot_id=<%=lot%>"><%=name_lot%></a></p>
              </div>
            </td>
            <td rowspan="6" width="120">&nbsp; </td>
          </tr>
          <tr> 
            <td bgcolor="#FFCC99" colspan="2" height="24"> 
              <p><b>������� ���� ����:</b><font color="#3333FF"> <b><font color="#FF0000"><%=tek_price%></font> 
                <%=valuta%> </b></font><%=kvo_stavok==0?" ������ ���!":" ��� ������� "+kvo_stavok+" ������."%></p>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#EBF3F2" height="4" valign="top" colspan="2"> 
              <p>&nbsp;</p>
              </td>
          </tr>
          <tr> 
            <td> 
              <p>������ ������: <%=start_date%></p>
            </td>
            <td> 
              <p>��������� ������: <%=end_date%></p>
            </td>
          </tr>
          <tr>
            <td bgcolor="#EBF3F2"> 
              <p align="right"><b><font size="2" color="#FF0000">����� ���� ���������:</font></b></p>
            </td>
            <td bgcolor="#EBF3F2" valign="top"> 
              <p>&nbsp;
                <input type="text" name="exp_date" maxlength="10">
              </p>
            </td>
          </tr>
          <tr> 
            <td> 
              <p align="right"><b>���� ��������:</b></p>
            </td>
            <td> 
              <p align="left"> <font color="#3333FF"> 
                <input type="checkbox" name="is_del" value="1">
                <font size="3" color="#CC0000"><b>��, � ������������� ���� �������� 
                ���! 
                <input type="submit" name="Submit" value="�������� ���">
                </b></font></font></p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <p align="center">&nbsp; </p>
  <p align="center">
    <%}else {%>
  </p>
  
</form>
<div align="right">
  <div align="center">
    <p><font size="3" color="#0000FF"><b>��� �������. </b></font></p>
  </div>
  <p align="center">&nbsp; </p>
  <p align="center"> 
    <%}%>
  </p>
</div>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#330000">
  <tr bgcolor="#FFCC99" bordercolor="#FFCC99"> 
    <td width="150"> 
      <h1>&nbsp;</h1>
    </td>
    <td colspan="2"> 
      <p><b><a href="auction.asp"><font face="Verdana, Arial, Helvetica, sans-serif">���� 
        �������</font></a> &gt;&gt; <font color="#6699FF"><a href="auction.asp?subj_id=<%=subj%>"> 
        <%=subj_name%></a></font></b></p>
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
      <h1 align="center"><font size="1">���������������� � ������ <a href="http://www.rusintel.ru/" target="_blank">��� 
        ��������</a> &copy; 2001-2002</font></h1>
    </td>
  </tr>
</table>
</body>
</html>
