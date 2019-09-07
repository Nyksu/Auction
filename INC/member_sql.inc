<%
var member_insert="insert into MEMBER (ID, NAME,  PSW, E_MAIL, PHONE, NIK_NAME, AUCTION_SUBJ_ID, ADM, KEYWORD, STATE, CITY, REG_DATE)"+
" values(%ID, '%NAME', '%PSW', '%EMAIL', '%PHONE', '%NIK', null, 0, '%CW', -1, '%CITY', 'TODAY')"

%>