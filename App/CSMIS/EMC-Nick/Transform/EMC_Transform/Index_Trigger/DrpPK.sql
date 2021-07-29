/*
DROP 轉檔會影響到的 PK
@C:\Transform\EMC_Transform\Index_Trigger\DrpPK.sql

Date: 2002.11.23
by: Jackal
*/

ALTER TABLE CD004 DROP CONSTRAINT PK_CD004;
ALTER TABLE CD019 DROP CONSTRAINT PK_CD019;
ALTER TABLE CD005 DROP CONSTRAINT PK_CD005;
ALTER TABLE CD031 DROP CONSTRAINT PK_CD031;
ALTER TABLE CD034 DROP CONSTRAINT PK_CD034;
ALTER TABLE CD007 DROP CONSTRAINT PK_CD007;
ALTER TABLE CD020 DROP CONSTRAINT PK_CD020;
ALTER TABLE CD021 DROP CONSTRAINT PK_CD021;
ALTER TABLE CD009 DROP CONSTRAINT PK_CD009;

alter table CD021 drop constraint UK_CD021_Description;

/*
alter table CD004 disable primary key cascade;
alter table CD019 disable primary key cascade;
alter table CD005 disable primary key cascade;
alter table CD031 disable primary key cascade;
alter table CD034 disable primary key cascade;
alter table CD007 disable primary key cascade;
alter table CD020 disable primary key cascade;
alter table CD021 disable primary key cascade;
alter table CD009 disable primary key cascade;

alter table CD021 disable validate constraint UK_CD021_Description;
alter table CD031 disable validate constraint UK_CD031_Description;
alter table CD034 disable validate constraint UK_CD034_Description;

*/
