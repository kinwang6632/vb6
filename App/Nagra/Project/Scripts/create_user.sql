/*
  ����: �إ� EMC ��Ʈw,CA�� User script files
	Use 'SYSTEM' to login SQL*plus and run this file
  �ɦW: create_user.sql
  Date: 2002.07.31
*/


-- 1. �إ� database user
DROP USER "COM";

CREATE USER "COM"  PROFILE "DEFAULT" IDENTIFIED BY "COM" 
    DEFAULT 
    TABLESPACE "COM" TEMPORARY 
    TABLESPACE "COM_TMP" ACCOUNT UNLOCK;

GRANT "CONNECT" TO "COM";
GRANT "RESOURCE" TO "COM";
GRANT UNLIMITED TABLESPACE TO "COM";
GRANT create table to "COM";
GRANT create view to "COM";
GRANT create sequence to "COM";

GRANT CREATE ANY TRIGGER TO "COM";
GRANT CREATE SNAPSHOT TO "COM";
GRANT GLOBAL QUERY REWRITE TO "COM";
GRANT QUERY REWRITE TO "COM";

ALTER USER "COM" DEFAULT ROLE ALL;


-- 2. grant �v�� => �n�� system�Ӱ���H�U�����O
GRANT ALTER ANY TRIGGER TO "COM";
