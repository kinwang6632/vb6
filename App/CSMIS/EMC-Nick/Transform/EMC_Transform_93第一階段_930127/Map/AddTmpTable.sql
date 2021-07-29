Prompt ***** 92.12.27 建立新舊代碼對照 Table 及 備份舊代碼檔 *****
--*******************************************************************
--*******************************************************************
Prompt 建立各資料檔轉檔完成Log檔 Table: TransformLog_1
DROP TABLE TransformLog_1 CASCADE CONSTRAINT;
create table TransformLog_1 (
	TableName varchar2(20),
	Description varchar2(1000),
	OkCounts number(9),
	ErrorCounts number(9),
	StartTime Date,
	StopTime Date,
	AllTime varchar2(100)
        );


Prompt 建立Exception Log檔 Table: TransformLog_2
DROP TABLE TransformLog_2 CASCADE CONSTRAINT;
create table TransformLog_2 (
	Type number(2),
	Description varchar2(1000),
	TableName varchar2(20)
        );

--Type=1 -[對照表有誤] 
--Type=2 -[其他錯誤] 



--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-媒體介紹類別代碼檔(CD009)之新舊代碼對照 Table: CD009_Map
DROP TABLE CD009_Map CASCADE CONSTRAINT;
create table CD009_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD009 舊代碼 Table *****
drop table CD009_Old cascade constraint;
create table CD009_Old as (select * from CD009);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-註銷原因代碼檔(CD012)之新舊代碼對照 Table: CD012_Map
DROP TABLE CD012_Map CASCADE CONSTRAINT;
create table CD012_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD012 舊代碼 Table *****
drop table CD012_Old cascade constraint;
create table CD012_Old as (select * from CD012);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-未繳費原因代碼檔(CD013)之新舊代碼對照 Table: CD013_Map
DROP TABLE CD013_Map CASCADE CONSTRAINT;
create table CD013_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD013 舊代碼 Table *****
drop table CD013_Old cascade constraint;
create table CD013_Old as (select * from CD013);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-退單原因代碼檔(CD015)之新舊代碼對照 Table: CD015_Map
DROP TABLE CD015_Map CASCADE CONSTRAINT;
create table CD015_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD015 舊代碼 Table *****
drop table CD015_Old cascade constraint;
create table CD015_Old as (select * from CD015);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-短收原因代碼檔(CD016)之新舊代碼對照 Table: CD016_Map
DROP TABLE CD016_Map CASCADE CONSTRAINT;
create table CD016_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD016 舊代碼 Table *****
drop table CD016_Old cascade constraint;
create table CD016_Old as (select * from CD016);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-付費意願代碼檔(CD020)之新舊代碼對照 Table: CD020_Map
DROP TABLE CD020_Map CASCADE CONSTRAINT;
create table CD020_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD020 舊代碼 Table *****
drop table CD020_Old cascade constraint;
create table CD020_Old as (select * from CD020);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-建築型態代碼檔(CD021)之新舊代碼對照 Table: CD021_Map
DROP TABLE CD021_Map CASCADE CONSTRAINT;
create table CD021_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD021 舊代碼 Table *****
drop table CD021_Old cascade constraint;
create table CD021_Old as (select * from CD021);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-服務滿意度代碼檔(CD026)之新舊代碼對照 Table: CD026_Map
DROP TABLE CD026_Map CASCADE CONSTRAINT;
create table CD026_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD026 舊代碼 Table *****
drop table CD026_Old cascade constraint;
create table CD026_Old as (select * from CD026);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-作廢原因代碼檔(CD051)之新舊代碼對照 Table: CD051_Map
DROP TABLE CD051_Map CASCADE CONSTRAINT;
create table CD051_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD051 舊代碼 Table *****
drop table CD051_Old cascade constraint;
create table CD051_Old as (select * from CD051);








Prompt OK
exit




