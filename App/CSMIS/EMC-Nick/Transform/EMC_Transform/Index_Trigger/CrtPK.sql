/*
Create 轉檔會影響到的 PK
@C:\Transform\EMC_Transform\Index_Trigger\CrtPK.sql

Date: 2002.11.23
by: Jackal
*/


ALTER TABLE CD004 ADD CONSTRAINT PK_CD004 PRIMARY KEY (CodeNo);
ALTER TABLE CD019 ADD CONSTRAINT PK_CD019 PRIMARY KEY (CodeNo);
ALTER TABLE CD005 ADD CONSTRAINT PK_CD005 PRIMARY KEY (CodeNo);
ALTER TABLE CD031 ADD CONSTRAINT PK_CD031 PRIMARY KEY (CodeNo);
ALTER TABLE CD034 ADD CONSTRAINT PK_CD034 PRIMARY KEY (CodeNo);
ALTER TABLE CD007 ADD CONSTRAINT PK_CD007 PRIMARY KEY (CodeNo);
ALTER TABLE CD020 ADD CONSTRAINT PK_CD020 PRIMARY KEY (CodeNo);
ALTER TABLE CD021 ADD CONSTRAINT PK_CD021 PRIMARY KEY (CodeNo);
ALTER TABLE CD009 ADD CONSTRAINT PK_CD009 PRIMARY KEY (CodeNo);

alter table CD021 add constraint UK_CD021_Description unique (description);

/*
alter table CD004 enable primary key cascade;
alter table CD019 enable primary key cascade;
alter table CD005 enable primary key cascade;
alter table CD031 enable primary key cascade;
alter table CD034 enable primary key cascade;
alter table CD007 enable primary key cascade;
alter table CD020 enable primary key cascade;
alter table CD021 enable primary key cascade;
alter table CD009 enable primary key cascade;

alter table CD021 enable validate constraint UK_CD021_Description;
alter table CD031 enable validate constraint UK_CD031_Description;
alter table CD034 enable validate constraint UK_CD034_Description;
*/

