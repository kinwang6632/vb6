rem ***** 91.11.21 建立新舊代碼對照 Table 及 備份舊代碼檔 *****
Sqlplusw.exe kp/kp@emc @C:\Transform\EMC_Transform\AddTmpTable.sql


rem ***** Load 轉換代碼之相關資料 *****
rem del *.log
rem del *.bad

rem 1. Load 客戶類別代碼檔(CD004)之新舊對照表: CD004_Map
sqlldr userid=kp/kp@emc control=CD004_Map.ctl  log=CD004_Map.log

rem 2. Load 客戶類別代碼檔(CD004)之新代碼表: CD004_New
sqlldr userid=kp/kp@emc control=CD004_New.ctl  log=CD004_New.log



rem 3. Load 收費項目代碼檔(CD019)之新舊對照表: CD019_Map
sqlldr userid=kp/kp@emc control=CD019_Map.ctl  log=CD019_Map.log

rem 4. Load 收費項目代碼檔(CD019)之新代碼表: CD019_New
sqlldr userid=kp/kp@emc control=CD019_New.ctl  log=CD019_New.log



rem 5. Load 裝機類別代碼檔(CD005)之新舊對照表: CD005_Map
sqlldr userid=kp/kp@emc control=CD005_Map.ctl  log=CD005_Map.log

rem 6. Load 裝機類別代碼檔(CD005)之新代碼表: CD005_New
sqlldr userid=kp/kp@emc control=CD005_New.ctl  log=CD005_New.log



rem 7. Load 收費方式代碼檔(CD031)之新舊對照表: CD031_Map
rem sqlldr userid=kp/kp@emc control=CD031_Map.ctl  log=CD031_Map.log

rem 8. Load 收費方式代碼檔(CD031)之新代碼表: CD031_New
rem sqlldr userid=kp/kp@emc control=CD031_New.ctl  log=CD031_New.log



rem 9. Load 買賣方式代碼檔(CD034)之新舊對照表: CD034_Map
rem sqlldr userid=kp/kp@emc control=CD034_Map.ctl  log=CD034_Map.log

rem 10. Load 買賣方式代碼檔(CD034)之新代碼表: CD034_New
rem sqlldr userid=kp/kp@emc control=CD034_New.ctl  log=CD034_New.log



rem 11. Load 停拆移機類別代碼檔(CD007)之新舊對照表: CD007_Map
sqlldr userid=kp/kp@emc control=CD007_Map.ctl  log=CD007_Map.log

rem 12. Load 停拆移機類別代碼檔(CD007)之新代碼表: CD007_New
sqlldr userid=kp/kp@emc control=CD007_New.ctl  log=CD007_New.log



rem 13. Load 付費意願代碼檔(CD020)之新舊對照表: CD020_Map
sqlldr userid=kp/kp@emc control=CD020_Map.ctl  log=CD020_Map.log

rem 14. Load 付費意願代碼檔(CD020)之新代碼表: CD020_New
sqlldr userid=kp/kp@emc control=CD020_New.ctl  log=CD020_New.log



rem 15. Load 建築型態代碼檔(CD021)之新舊對照表: CD021_Map
sqlldr userid=kp/kp@emc control=CD021_Map.ctl  log=CD021_Map.log

rem 16. Load 建築型態代碼檔(CD021)之新代碼表: CD021_New
sqlldr userid=kp/kp@emc control=CD021_New.ctl  log=CD021_New.log



rem 17. Load 媒體介紹類別代碼檔(CD009)之新舊對照表: CD009_Map
sqlldr userid=kp/kp@emc control=CD009_Map.ctl  log=CD009_Map.log

rem 18. Load 媒體介紹類別代碼檔(CD009)之新代碼表: CD009_New
sqlldr userid=kp/kp@emc control=CD009_New.ctl  log=CD009_New.log















