<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\member_sql.inc" -->
<!-- #include file="inc\email.inc" -->


<%
var ErrorMsg=""
var cw=""
var isSending=false
var city=""

var eml=Server.CreateObject("JMail.Message")
eml.Logging=false
eml.From="auction@72rus.ru"
eml.Subject="��� ��������� ��������� �������� �� ������� auction.72rus.ru"
eml.Charset=characterset
eml.ContentTransferEncoding = "base64"
eml.FromName="��������� �������"

isFirst=String(Request.Form("Submit"))=="undefined"
ShowForm=true
TestKey=(String(Session("TK"))!="undefined")
if(!isFirst){
	
		//-------------input validation-----------
		Nik=TextFormData(Request.Form("nik"),"")
		Name=TextFormData(Request.Form("Name"),"")
		Email=TextFormData(Request.Form("e_mail"),"")
		city=TextFormData(Request.Form("city"),"")
		if(TestKey){
			Keyword=TextFormData(Request.Form("keyword"),"")
			if(Keyword.length<20){ErrorMsg+="������������ ���������� ���� '��� ���������'.<br>"}
			else{
				Records.Source="Select Count(ID) as ID from MEMBER where KEYWORD='"+Keyword+"' and NIK_NAME='"+Nik+"' and STATE=-1"
				Records.Open()
				if (Records("ID").Value==0){ErrorMsg+="�������� '��� ���������'.<br>"}
				Records.Close()
			}
			ShowForm=false
		}
		else{

		Pass=TextFormData(Request.Form("pass"),"")
		Pass2=TextFormData(Request.Form("pass2"),"")
		Phone=TextFormData(Request.Form("phone"),"")
		if(Name.length<6){ErrorMsg+="������������ ���������� ���� '�.�.�.'.<br>"}
		if(Nik.length<3){ErrorMsg+="'���������' ������ ���� ������ 2-� ��������.<br>"}
		if((Phone!="")&&(Phone.length<5)){ErrorMsg+="������������ ���������� �����.<br>"}
		if(Pass.length<5){ErrorMsg+="���� '������' ������ �������� �� ����� ��� �� 5 ��������.<br>"}
		if(Pass != Pass2){ErrorMsg+="�� ��������� ������! ���������� ������� �� ������.<br>"}
		if (city.length<2) {ErrorMsg+="������� �������� �������� ������.<br>"}

		if (!/(\w+)@((\w+).)*(\w+)$/.test(Email)) {ErrorMsg=ErrorMsg+"�������� ������ ���� 'E-mail'.<br>"}

		Records.Source="Select Count(ID) as ID from MEMBER where NIK_NAME='"+Nik+"'"
		Records.Open()
		if (Records("ID").Value>0){ErrorMsg+="����� ��� ������������ ��� ����������.<br>"}
		Records.Close()

		}

		if ((ErrorMsg=="")&&(TestKey)){
			sql="Update MEMBER set STATE=0 where KEYWORD='"+Keyword+"' and NIK_NAME='"+Nik+"'"
			try{
				Connect.Execute(sql)
				Session("TK")="undefined"
	 			TestKey=false
			}
			catch(e){
				ErrorMsg+=ListAdoErrors()
			}
		}

		if ((ErrorMsg=="")&&(ShowForm)){
			id=NextID("UNIVERSAL")
			cw=String(Math.random()).substr(3,20)
			if(cw.length<20){cw+=String(Math.random()).substr(3,(20-cw.length))}
			sql=member_insert
			sql=sql.replace("%ID",id)
			sql=sql.replace("%NAME",Name)  
			sql=sql.replace("%NIK",Nik)  
			sql=sql.replace("%PHONE",Phone)
			sql=sql.replace("%EMAIL",Email)
			sql=sql.replace("%CITY",city)
			sql=sql.replace("%PSW",Pass)
			sql=sql.replace("%CW",String(cw))
			eml.AddRecipient(Email)
			eml.AppendText(Name+", �� ������������������ � �������� ������� ��������� �������!")
			eml.AppendText("\n ��� ������������: "+Nik)
			eml.AppendText("\n  ������: "+Pass)
			eml.AppendText("\n \n ��������!! ���� �������� ��� ���������, ������� ���������� ������������ ���� ������� ������!")
			eml.AppendText("\n ��� ��������� : "+String(cw))
			eml.AppendText("\n \n ������� auction.72rus.ru")
			try{
				isSending=eml.Send(servsmtp)
	 			if (!isSending) {ErrorMsg+="�������� � ������.<br>"}
				else{
				Connect.Execute(sql)
	 			ShowForm=false
				Session("TK")=true
				TestKey=true
				}
			}
			catch(e){
				ErrorMsg+=ListAdoErrors()
				if (!isSending) {ErrorMsg+="�������� � ������.<br>"}
			}
		} 
		
	
}
%>

<html>
<head>
<title>��������� �������</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<style type="text/css">
<!--
p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 10pt; font-weight: 400; color: #003366; margin:  3px}
a:hover {color:#FF3333}
a:visited {  font-family: Arial, Helvetica, sans-serif; color: #666666}
h1 {  font: 12pt Verdana, Arial, Helvetica, sans-serif; color: #000099}
a:link {  font: 9pt/9pt Arial, Helvetica, sans-serif; color: #6666FF}
-->
</style>
<style type="text/css">
<!--
p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 10pt; font-weight: 400; color: #000000; margin:  3px}
a:hover {color:#FF3333; font: 9pt/9pt Arial, Helvetica, sans-serif}
a:visited {  font-family: Arial, Helvetica, sans-serif; color: #990000; font-size: 9pt; line-height: 9pt}
h1 {  font: 10pt/normal Verdana, Arial, Helvetica, sans-serif; color: #000000; text-decoration: blink; border-width: 1px auto medium medium; border-color: #CCCCFF black black; background: repeat-y; margin: 3px 3px}
a:link {  font: 9pt/9pt Arial, Helvetica, sans-serif; color: #000000}
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
      size=6><bl><font color="#FFCC99" face="Times New Roman, Times, serif">��������� 
        �������</font><font color="#FFFFFF"><br>
        <font face="Arial, Helvetica, sans-serif" 
      size=6 color="#FFFFFF"><b><font size="1" color="#CC9900">������� �������� 
        ����� / </font><font color=#FFFFFF face="Verdana, Arial, Helvetica, sans-serif" 
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
        </font></b></font></font></bl></font></div>
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
      <p>:: <a href="auction.asp">�������</a> | <a href="auction_rules.html">������� 
        ���������</a> | <a href="area.asp">������ �������</a> | <a href="reconnect.asp">����� 
        ��� ������ ������</a> ::</p>
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
      <h1> :: ����������� ��������� �������� ::</h1>
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
    <td width="150" bordercolor="#FFCC99" bgcolor="#FFCC99" background="Images/line_back.gif"> 
      <p>&nbsp;</p>
      </td>
    <td bgcolor="#FFFFFF" bordercolor="#FFFFFF">
      <table width="100%" border="0">
        <tr valign="top" align="left"> 
          <td height="485" > 
            <h1>
              <%if(ErrorMsg!=""){%>
            </h1>
            <h2> 
              <p> <font color="#FF3300" size="2"><b>������!</b></font> <br>
                <%=ErrorMsg%></p>
            </h2>
            <%}%>
            <form name="reg_auc+mem" method="post" action="reg_mem_auction.asp">
              <%if((ShowForm) || (TestKey)){%>
			  <p><b>����� ����������� ������ ����� ����������� ����������� � </b><a href="auction_rules.html">��������� 
              ���������� ��������</a>!</p>
              <p><b>�������� ����� �� ������������ �� ����� ��������� ��������! 
                !</b></p>
              <p>&nbsp;</p>
			  <table width="100%" border="0" cellspacing="4" cellpadding="0">
                <%if (!TestKey){%>
                <tr valign="middle"> 
                  <td width="30%" align="right"> 
                    <p><b>�.�.�. (��� �������� ��. ����)</b></p>
                  </td>
                  <td width="70%"> 
                    <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" maxlength="80" size="40">
                    * </td>
                </tr>
                <tr valign="middle"> 
                  <td width="30%" align="right"> 
                    <p><b>��������� (�����)</b></p>
                  </td>
                  <td width="70%"> 
                    <input type="text" name="nik" value="<%=isFirst?"":Request.Form("nik")%>" size="20" maxlength="20">
                    * </td>
                </tr>
				<tr valign="middle"> 
                  <td width="30%" align="right"> 
                    <p><b>������</b></p>
                  </td>
                  <td width="70%"> 
                    <input type="password" name="pass" size="20" maxlength="20">
                    * </td>
                </tr>
                <tr valign="middle"> 
                  <td width="30%" align="right"> 
                    <p><b>������ (��������)</b></p>
                  </td>
                  <td width="70%"> 
                    <input type="password" name="pass2" size="20" maxlength="20">
                    * </td>
                </tr>
                <tr valign="middle"> 
                  <td width="30%" align="right"> 
                    <p><b>�����</b></p>
                  </td>
                  <td width="70%"> 
                    <input type="text" name="city" maxlength="80" size="30" value="<%=isFirst?"":Request.Form("city")%>">
                    * </td>
                </tr>
                <tr valign="middle"> 
                  <td width="30%" align="right"> 
                    <p><b>�������</b></p>
                  </td>
                  <td width="70%"> 
                    <input type="text" name="phone" value="<%=isFirst?"":Request.Form("phone")%>" size="40" maxlength="40">
                  </td>
                </tr>
                <tr valign="middle"> 
                  <td width="30%" align="right"> 
                    <p><b>E-mail</b></p>
                  </td>
                  <td width="70%"> 
                    <input type="text" name="e_mail" value="<%=isFirst?"":Request.Form("e_mail")%>" size="40" maxlength="40">
                    *</td>
                </tr>
                <%} else {%>
                <tr> 
                  <td width="30%" height="5"> 
                    <p align="right"><b>��� ��������� </b></p>
                    <p align="right">(������ ��� �� ��������� E-mail)</p>
                  </td>
                  <td width="70%" height="5"> 
                    <p>
                      <input type="text" name="keyword" size="20" maxlength="20">
                      <b><font color="#FF0000">��������!!</font></b> ���� ��� 
                      �� ������ ��������� ����� ��������� �� ��� �������� ���� 
                      (��������� �����) ���� ���������! ���� �� �� ������������ 
                      ��������, �� ������ ������ ���� ��� ��� ��������� ��������� 
                      �������� � ������ ����������� (����� � ������ �������). 
                      <input type="hidden" name="name2" value="<%=Request.Form("name")%>">
                      <input type="hidden" name="nik" value="<%=Request.Form("nik")%>">
                      <input type="hidden" name="phone2" value="<%=Request.Form("phone")%>">
                      <input type="hidden" name="e_mail" value="<%=Request.Form("e_mail")%>">
                    </p>
                  </td>
                </tr>
                <%}%>
                <tr> 
                  <td width="30%" height="22"> 
                    <p>&nbsp;</p>
                  </td>
                  <td width="70%" height="22"> 
                    <p>&nbsp;</p>
                    <p>* - ������������ ����</p>
                  </td>
                </tr>
              </table>
              <div align="center"> 
                <input type="submit" name="Submit" value="����������� �����������">
              </div>
            </form>
            <p align="center"><a href="auction.asp">��������� � �������</a></p>
            <hr>
            <%}
else{%>
            <center>
              <h2><font color="#3333FF">����������� ������ �������!</font></h2>
              <p>&nbsp;</p>
              <p><a href="area.asp">����� � ������ �������!</a></p>
              <p>&nbsp;</p>
              <p><font face="Arial, Helvetica, sans-serif"><a href="auction.asp">��������� 
                �������</a></font></p>
              <p>&nbsp; </p>
              <p>&nbsp; </p>
            </center>
            <%}%>
          </td>
        </tr>
      </table>
      <p>&nbsp;</p>
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
      <h1 align="center"><font size="1">���������������� � ������ <a href="http://www.rusintel.ru/" target="_blank">��� 
        ��������</a> &copy; 2001-2002</font></h1>
    </td>
  </tr>
</table>
</body>
</html>
