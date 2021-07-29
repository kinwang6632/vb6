
rem ***** Load EMC新增代碼 *****
rem del *.log
rem del *.bad

rem 0.刪除相關Table
Sqlplusw.exe v30/v30@emc @C:\Transform\EMC_Transform2\DeleteTable.sql

rem 1. Load 收費項目代碼檔(CD019)之新代碼 : CD019
sqlldr userid=v30/v30@emc control=CD019_Map.ctl  log=CD019_Map.log

rem 2. Load 頻道編號代碼檔(CD024)之新代碼 : CD024
sqlldr userid=v30/v30@emc control=CD024_Map.ctl  log=CD024_Map.log

rem 3. Load 收費頻道細項檔(CD019A)之新代碼 : CD019A
sqlldr userid=v30/v30@emc control=CD019A_Map.ctl  log=CD019A_Map.log

rem 4. Load 裝機類別代碼檔(CD005)之新代碼 : CD005
sqlldr userid=v30/v30@emc control=CD005_Map.ctl  log=CD005_Map.log

rem 5. Load 停拆移機類別代碼檔(CD007)之新代碼 : CD007
sqlldr userid=v30/v30@emc control=CD007_Map.ctl  log=CD007_Map.log

rem 6. Load 品名編號代碼檔(CD022)之新代碼 : CD022
sqlldr userid=v30/v30@emc control=CD022_Map.ctl  log=CD022_Map.log

rem 7. Load 設備買賣付款方式對應檔(CD022A)之新代碼 : CD022A
sqlldr userid=v30/v30@emc control=CD022A_Map.ctl  log=CD022A_Map.log

rem 8. Load 電話申告類別代碼檔(CD008)之新代碼 : CD008
sqlldr userid=v30/v30@emc control=CD008_Map.ctl  log=CD008_Map.log

rem 9. Load 申告內容代碼檔(CD008A)之新代碼 : CD008A
sqlldr userid=v30/v30@emc control=CD008A_Map.ctl  log=CD008A_Map.log

rem 10. Load 申告類別之申告內容(CD008C)之新代碼 : CD008C
sqlldr userid=v30/v30@emc control=CD008C_Map.ctl  log=CD008C_Map.log

rem 11. Load 故障原因代碼檔(CD011)之新代碼 : CD011
sqlldr userid=v30/v30@emc control=CD011_Map.ctl  log=CD011_Map.log

rem 12. Load 設備裝置點代碼檔(CD056)之新代碼 : CD056
sqlldr userid=v30/v30@emc control=CD056_Map.ctl  log=CD056_Map.log

rem 13. Load 媒體介紹類別代碼檔(CD009)之新代碼 : CD009
sqlldr userid=v30/v30@emc control=CD009_Map.ctl  log=CD009_Map.log

rem 14. Load 付款種類代碼檔(CD032)之新代碼 : CD032
sqlldr userid=v30/v30@emc control=CD032_Map.ctl  log=CD032_Map.log

rem 15. Load STB結清原因代碼(CD061)之新代碼 : CD061
sqlldr userid=v30/v30@emc control=CD061_Map.ctl  log=CD061_Map.log

rem 16. Load 頻道結清原因代碼(CD060)之新代碼 : CD060
sqlldr userid=v30/v30@emc control=CD060_Map.ctl  log=CD060_Map.log

rem 17. Load STB配件代碼檔(CD028)之新代碼 : CD028
sqlldr userid=v30/v30@emc control=CD028_Map.ctl  log=CD028_Map.log

rem 18. Load 訂單類別代碼檔(CD057)之新代碼 : CD057
sqlldr userid=v30/v30@emc control=CD057_Map.ctl  log=CD057_Map.log







