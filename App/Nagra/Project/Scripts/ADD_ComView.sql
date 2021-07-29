/*
	grant select on emcyms.cd024 to com;  
*/
set heading off
variable str varchar2(80)
exec :str := '<< 建立暫存檔 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool D:\App\Invoice\Log 
print str 

prompt CD024: 供倉管讀取的 view
drop view CD024;
create view CD024 as select CODENO , DESCRIPTION, CHANNELID from emcyms.cd024 where  STOPFLAG = 0;

/* 用法: select * from CD024 */

spool off
set heading on



