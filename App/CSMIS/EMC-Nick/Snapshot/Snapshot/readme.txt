�Y�n�N MATERIALIZED VIEW ��b���P�� TABLESPACE, ���O�p�U:
CREATE MATERIALIZED all_emps TABLESPACE snap_ts 
USING INDEX TABLESPACE ind_snap_ts 
AS SELECT * FROM scott.emp;  


�ݫإߪ����G, ���O�p�U:
desc mlog$_tableName; => �� master site ���ͪ� trigger
desc tlog$_tableName; => �� target site ���ͪ� trigger


�Y�n�ϥΤw�g�s�b��table�� MATERIALIZED VIEW, ���O�p�U:
CREATE MATERIALIZED VIEW v30.send_nagra
ON PREBUILT TABLE 
REFRESH FAST
ENABLE QUERY REWRITE
AS 
SELECT * 
FROM V30.send_nagra@RDSERVER.US.ORACLE.COM;

