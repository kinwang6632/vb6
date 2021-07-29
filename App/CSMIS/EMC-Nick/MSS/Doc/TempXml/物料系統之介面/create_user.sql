/*
  說明: 建立 EMC 資料庫,物料系統介面之 User script files
	Use 'SYSTEM' to login SQL*plus and run this file
  檔名: create_user.sql
  Date: 2002.07.31
*/


-- 1. 建立 database user
DROP USER "EMCMSS";

CREATE USER "EMCMSS"  PROFILE "DEFAULT" IDENTIFIED BY "emcmss" 
    DEFAULT 
    TABLESPACE "COM" TEMPORARY 
    TABLESPACE "COM_TMP" ACCOUNT UNLOCK;

GRANT "CONNECT" TO "EMCMSS";
GRANT "RESOURCE" TO "EMCMSS";
GRANT UNLIMITED TABLESPACE TO "EMCMSS";
ALTER USER "EMCMSS" DEFAULT ROLE ALL;

grant create table to "EMCMSS";
grant create view to "EMCMSS";
grant create sequence to "EMCMSS";

