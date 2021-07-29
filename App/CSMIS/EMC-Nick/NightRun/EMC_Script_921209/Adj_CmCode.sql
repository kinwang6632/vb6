/*
  調整EMC各家的收費方式代碼, 以及有用到收費方式的資料檔. 

  調整方式:
  < 前置作業 >
	1. 做出每一家的收費方式代碼新舊值對照表
	2. 依據對照表, 寫出該家的的轉換script file

  < 現場作業 >
	1. login每一家
	2. compile轉代碼資料函式
	3. 執行該家的代碼轉換程式

  注意事項:
  . 因為可能調整大量資料, 故可指定rollback segment
*/

set serveroutput on
set heading off
variable str varchar2(80)
exec :str := '<< 調整EMC各家的收費方式代碼 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool c:\gird\Log\Adj_CmCode.log
print str


-- 1. 調整EMCYMS
conn EMCYMS/EMCYMS@????
@c:\gird\v300\script\SP_UpdCode
exec SP_UpdCode('xxx', a, b);
set transaction use rollback segment rbS27;


spool off
set heading on
