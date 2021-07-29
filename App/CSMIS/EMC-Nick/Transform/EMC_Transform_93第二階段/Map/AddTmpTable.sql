Prompt ***** 93.02.17 建立新舊代碼對照 Table 及 備份舊代碼檔 *****
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
Prompt 建立-客戶類別代碼檔(CD004)之新舊代碼對照 Table: CD004_Map
DROP TABLE CD004_Map CASCADE CONSTRAINT;
create table CD004_Map (
	NewCode number(5),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD004 舊代碼 Table *****
drop table CD004_Old cascade constraint;
create table CD004_Old as (select * from CD004);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-裝機類別代碼檔(CD005)之新舊代碼對照 Table: CD005_Map
DROP TABLE CD005_Map CASCADE CONSTRAINT;
create table CD005_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD005 舊代碼 Table *****
drop table CD005_Old cascade constraint;
create table CD005_Old as (select * from CD005);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-維修申告類別代碼檔(CD006)之新舊代碼對照 Table: CD006_Map
DROP TABLE CD006_Map CASCADE CONSTRAINT;
create table CD006_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD006 舊代碼 Table *****
drop table CD006_Old cascade constraint;
create table CD006_Old as (select * from CD006);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-停拆移機類別代碼檔(CD007)之新舊代碼對照 Table: CD007_Map
DROP TABLE CD007_Map CASCADE CONSTRAINT;
create table CD007_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD007 舊代碼 Table *****
drop table CD007_Old cascade constraint;
create table CD007_Old as (select * from CD007);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-電話申告類別代碼檔(CD008)之新舊代碼對照 Table: CD008_Map
DROP TABLE CD008_Map CASCADE CONSTRAINT;
create table CD008_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD008 舊代碼 Table *****
drop table CD008_Old cascade constraint;
create table CD008_Old as (select * from CD008);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-申告內容代碼檔(CD008A)之新舊代碼對照 Table: CD008A_Map
DROP TABLE CD008A_Map CASCADE CONSTRAINT;
create table CD008A_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD008A 舊代碼 Table *****
drop table CD008A_Old cascade constraint;
create table CD008A_Old as (select * from CD008A);


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
Prompt 建立-故障原因代碼檔(CD011)之新舊代碼對照 Table: CD011_Map
DROP TABLE CD011_Map CASCADE CONSTRAINT;
create table CD011_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD011 舊代碼 Table *****
drop table CD011_Old cascade constraint;
create table CD011_Old as (select * from CD011);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-故障原因代碼檔2(CD011B)之新舊代碼對照 Table: CD011B_Map
DROP TABLE CD011B_Map CASCADE CONSTRAINT;
create table CD011B_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD011B 舊代碼 Table *****
drop table CD011B_Old cascade constraint;
create table CD011B_Old as (select * from CD011B);


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
Prompt 建立-停拆機原因代碼檔(CD014)之新舊代碼對照 Table: CD014_Map
DROP TABLE CD014_Map CASCADE CONSTRAINT;
create table CD014_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD014 舊代碼 Table *****
drop table CD014_Old cascade constraint;
create table CD014_Old as (select * from CD014);


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
Prompt 建立-銀行代碼檔(CD018)之新舊代碼對照 Table: CD018_Map
DROP TABLE CD018_Map CASCADE CONSTRAINT;
create table CD018_Map (
	NewCode number(3),
	NewName varchar2(24),
	OldCode number(3),
	OldName varchar2(24),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD018 舊代碼 Table *****
drop table CD018_Old cascade constraint;
create table CD018_Old as (select * from CD018);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-收費項目代碼檔(CD019)之新舊代碼對照 Table: CD019_Map
DROP TABLE CD019_Map CASCADE CONSTRAINT;
create table CD019_Map (
	NewCode number(5),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD019 舊代碼 Table *****
drop table CD019_Old cascade constraint;
create table CD019_Old as (select * from CD019);


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
Prompt 建立-品名編號代碼檔(CD022)之新舊代碼對照 Table: CD022_Map
DROP TABLE CD022_Map CASCADE CONSTRAINT;
create table CD022_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD022 舊代碼 Table *****
drop table CD022_Old cascade constraint;
create table CD022_Old as (select * from CD022);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-客戶來源代碼檔(CD025)之新舊代碼對照 Table: CD025_Map
DROP TABLE CD025_Map CASCADE CONSTRAINT;
create table CD025_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD025 舊代碼 Table *****
drop table CD025_Old cascade constraint;
create table CD025_Old as (select * from CD025);


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
Prompt 建立-消息來源檔(CD049)之新舊代碼對照 Table: CD049_Map
DROP TABLE CD049_Map CASCADE CONSTRAINT;
create table CD049_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD049 舊代碼 Table *****
drop table CD049_Old cascade constraint;
create table CD049_Old as (select * from CD049);


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


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-管線類別代碼檔(CD055)之新舊代碼對照 Table: CD055_Map
DROP TABLE CD055_Map CASCADE CONSTRAINT;
create table CD055_Map (
	NewCode number(3),
	NewName varchar2(26),
	OldCode number(3),
	OldName varchar2(26),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD055 舊代碼 Table *****
drop table CD055_Old cascade constraint;
create table CD055_Old as (select * from CD055);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt 建立-商贈品代碼檔(CD059)之新舊代碼對照 Table: CD059_Map
DROP TABLE CD059_Map CASCADE CONSTRAINT;
create table CD059_Map (
	NewCode number(8),
	NewName varchar2(26),
	OldCode number(8),
	OldName varchar2(26),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** 備份 CD059 舊代碼 Table *****
drop table CD059_Old cascade constraint;
create table CD059_Old as (select * from CD059);





Prompt OK
exit




