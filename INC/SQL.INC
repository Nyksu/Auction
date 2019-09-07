<%
var select_goods="Select t1.*, t2.name as type_name, t3.KVO_TRAD as kvo, t4.city "+
"from goods t1, auction_type t2, GET_TRADE_COUNT_LIST(%SUBJ_ID) t3, member t4 "+
"where "+
"t1.state=0 "+
"and date_end>='TODAY' "+
"and t1.auction_type_id=t2.id "+
"and t1.auction_subj_id=%SUBJ_ID "+
"and t1.id=t3.id_goods "+
"and t1.member_id=t4.id"

var insert_lot="insert into goods (id, member_id, auction_subj_id, auction_type_id, name, start_price, date_begin, "+
"date_end, state, delta, price, valuta, z_price) "+
"values(%ID, %MID, %SUBJID, %TYPEID, '%NAME', %STARTPRICE, 'TODAY', '%DATEND', 0, %DD, %PRICE, '%VALUTA', %ZPRICE)"

var select_own_lots="Select t1.*, t2.name as subjname, t4.name as type_name"+
" from goods t1, auction_subj t2, auction_type t4"+
" where t1.member_id=%HOST_ID"+
" and t1.date_end>='TODAY'"+
" and t1.state=0 and t4.id=t1.auction_type_id"+
" and t1.auction_subj_id=t2.id"

var select_st_lots="Select t1.*, t2.name as subjname, t4.name as type_name, t5.member_id as last_mem"+
" from goods t1, auction_subj t2, auction_type t4, TRADE_HISTORY t5"+
" where t1.id=t5.goods_id and t5.member_id=%HOST_ID"+
" and t1.date_end>='TODAY'"+
" and t1.state=0 and t4.id=t1.auction_type_id"+
" and t1.auction_subj_id=t2.id and t5.state=0"

var do_rate="Select * from ADDRATE(%LOT,%RATE,%MEMID)"

var select_allfor_lots="Select t1.*, t2.name as subjname, t4.name as type_name"+
" from goods t1, auction_subj t2, auction_type t4"+
" where t1.id=%LOT_ID"+
" and t1.state=0 and t4.id=t1.auction_type_id"+
" and t1.auction_subj_id=t2.id"

var select1_allfor_lots="Select t1.*, t2.name as subjname, t4.name as type_name"+
" from goods t1, auction_subj t2, auction_type t4"+
" where t1.id=%LOT_ID"+
" and t4.id=t1.auction_type_id"+
" and t1.auction_subj_id=t2.id"

var del_lot="Select * from DELLOT(%LOT)"

var end_own_lots="Select t1.*, t2.name as subjname, t4.name as type_name"+
" from goods t1, auction_subj t2, auction_type t4"+
" where t1.member_id=%HOST_ID"+
" and (t1.state=1 or (t1.state=0 and t1.date_end<'TODAY'))"+
" and t4.id=t1.auction_type_id"+
" and t1.auction_subj_id=t2.id"

var end_st_lots="Select t1.*, t2.name as subjname, t4.name as type_name, t5.member_id as last_mem"+
" from goods t1, auction_subj t2, auction_type t4, TRADE_HISTORY t5"+
" where t1.id=t5.goods_id and t5.member_id=%HOST_ID"+
" and (t1.state=1 or (t1.state=0 and t1.date_end<'TODAY'))"+
" and t4.id=t1.auction_type_id and t5.state=0"+
" and t1.auction_subj_id=t2.id and t1.price=t5.price"

var winner_lot="Select t2.* from TRADE_HISTORY t1, member t2, goods t3 "+
"where t1.member_id=t2.id and "+
"t1.price=t3.price and t3.id=%LOT"

var set_lot_proc="Update goods set state=2 where id=%LOT"

var set_lot_prolong="Update goods set date_end='%DE' where id=%LOT"

var asklotadd="Insert into ask_goods (id,member_id,goods_id,name,ask_date) values(%ID, %MEM, %GD,'%NAME','TODAY')"

var select_new_goods="Select t1.*, t2.name as type_name, t4.nik_name "+
"from goods t1, auction_type t2, member t4 "+
"where "+
"t1.state=0 "+
"and date_begin>='TODAY'-3 "+
"and t1.auction_type_id=t2.id "+
"and t1.member_id=t4.id order by t1.id desc"

%>