set heading off
variable str varchar2(80)
exec :str := '<< 建立暫存檔 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool D:\App\EMC\EBT介接\Log
print str 


drop index I_EMC_EBT001_1;
	CREATE INDEX I_EMC_EBT001_1 ON EMC_EBT001 (EMCCUSTID, EBTCUSTID, EBTCONTRACTNO);

drop index I_SO001_1;
	CREATE INDEX I_SO001_1 ON SO001 (INSTADDRESS);

commit;
spool off
set heading on



