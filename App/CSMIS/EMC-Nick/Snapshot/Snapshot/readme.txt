若要將 MATERIALIZED VIEW 放在不同的 TABLESPACE, 指令如下:
CREATE MATERIALIZED all_emps TABLESPACE snap_ts 
USING INDEX TABLESPACE ind_snap_ts 
AS SELECT * FROM scott.emp;  


看建立的結果, 指令如下:
desc mlog$_tableName; => 看 master site 產生的 trigger
desc tlog$_tableName; => 看 target site 產生的 trigger


若要使用已經存在的table為 MATERIALIZED VIEW, 指令如下:
CREATE MATERIALIZED VIEW v30.send_nagra
ON PREBUILT TABLE 
REFRESH FAST
ENABLE QUERY REWRITE
AS 
SELECT * 
FROM V30.send_nagra@RDSERVER.US.ORACLE.COM;

