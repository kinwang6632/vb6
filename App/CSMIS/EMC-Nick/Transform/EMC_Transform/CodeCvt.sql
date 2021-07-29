-- 91.11.22
variable nn number
variable msg varchar2(200)

variable t1 varchar2(19)
variable t2 varchar2(19)
variable str varchar2(80)
variable tt number
variable hh number
variable mm number
variable ss number


set serveroutput on
set heading off

exec :t1 := to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
exec :str := '<< EMC : 正式更新各種代碼檔與資料檔 >> ' || :t1;
spool C:\Transform\EMC_Transform\CodeCvt.log
print str 


@C:\Transform\EMC_Transform\Index_Trigger\Drop.sql

-- 1. 指定RollBack segment 
-- 2. 更新代碼檔相關 Table 之欄位
set transaction use rollback segment rbS27;
@C:\Transform\EMC_Transform\SP_UpdData_CD004.SQL
exec SP_UpdData_CD004;
COMMIT;


set transaction use rollback segment rbS27;
@C:\Transform\EMC_Transform\SP_UpdData_CD019.SQL
exec SP_UpdData_CD019;
COMMIT;

set transaction use rollback segment rbS27;
@C:\Transform\EMC_Transform\SP_UpdData_CD005.SQL
exec SP_UpdData_CD005;
COMMIT;


set transaction use rollback segment rbS27;
@C:\Transform\EMC_Transform\SP_UpdData_CD007.SQL
exec SP_UpdData_CD007;
COMMIT;

set transaction use rollback segment rbS27;
@C:\Transform\EMC_Transform\SP_UpdData_CD020.SQL
exec SP_UpdData_CD020;
COMMIT;

set transaction use rollback segment rbS27;
@C:\Transform\EMC_Transform\SP_UpdData_CD021.SQL
exec SP_UpdData_CD021;
COMMIT;

set transaction use rollback segment rbS27;
@C:\Transform\EMC_Transform\SP_UpdData_CD009.SQL
exec SP_UpdData_CD009;
COMMIT;


@C:\Transform\EMC_Transform\Index_Trigger\Create.sql

exec :t2 := to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
exec :str := '<<轉檔完畢, 請檢查是否有錯誤 >> ' || :t2;
print str
exec :tt := TRUNC(86400*(to_date(:t2, 'YYYY/MM/DD HH24:MI:SS')-to_date(:t1, 'YYYY/MM/DD HH24:MI:SS')));
exec :hh := trunc(:tt/3600);
exec :ss := mod(:tt, 60);
exec :mm := (:tt - :hh*3600 - :ss) / 60;


exec :str := '合計執行時間: ' || :hh || '時' || :mm || '分' || :ss || '秒';
print str 

spool off
set heading on
EXIT
