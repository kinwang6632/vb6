Prompt ***** 93.02.17 �إ߷s�¥N�X��� Table �� �ƥ��¥N�X�� *****
--*******************************************************************
--*******************************************************************
Prompt �إߦU��������ɧ���Log�� Table: TransformLog_1
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


Prompt �إ�Exception Log�� Table: TransformLog_2
DROP TABLE TransformLog_2 CASCADE CONSTRAINT;
create table TransformLog_2 (
	Type number(2),
	Description varchar2(1000),
	TableName varchar2(20)
        );

--Type=1 -[��Ӫ��~] 
--Type=2 -[��L���~] 


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�Ȥ����O�N�X��(CD004)���s�¥N�X��� Table: CD004_Map
DROP TABLE CD004_Map CASCADE CONSTRAINT;
create table CD004_Map (
	NewCode number(5),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD004 �¥N�X Table *****
drop table CD004_Old cascade constraint;
create table CD004_Old as (select * from CD004);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�˾����O�N�X��(CD005)���s�¥N�X��� Table: CD005_Map
DROP TABLE CD005_Map CASCADE CONSTRAINT;
create table CD005_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD005 �¥N�X Table *****
drop table CD005_Old cascade constraint;
create table CD005_Old as (select * from CD005);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-���ץӧi���O�N�X��(CD006)���s�¥N�X��� Table: CD006_Map
DROP TABLE CD006_Map CASCADE CONSTRAINT;
create table CD006_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD006 �¥N�X Table *****
drop table CD006_Old cascade constraint;
create table CD006_Old as (select * from CD006);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-��������O�N�X��(CD007)���s�¥N�X��� Table: CD007_Map
DROP TABLE CD007_Map CASCADE CONSTRAINT;
create table CD007_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD007 �¥N�X Table *****
drop table CD007_Old cascade constraint;
create table CD007_Old as (select * from CD007);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�q�ܥӧi���O�N�X��(CD008)���s�¥N�X��� Table: CD008_Map
DROP TABLE CD008_Map CASCADE CONSTRAINT;
create table CD008_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD008 �¥N�X Table *****
drop table CD008_Old cascade constraint;
create table CD008_Old as (select * from CD008);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�ӧi���e�N�X��(CD008A)���s�¥N�X��� Table: CD008A_Map
DROP TABLE CD008A_Map CASCADE CONSTRAINT;
create table CD008A_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD008A �¥N�X Table *****
drop table CD008A_Old cascade constraint;
create table CD008A_Old as (select * from CD008A);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�C�餶�����O�N�X��(CD009)���s�¥N�X��� Table: CD009_Map
DROP TABLE CD009_Map CASCADE CONSTRAINT;
create table CD009_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD009 �¥N�X Table *****
drop table CD009_Old cascade constraint;
create table CD009_Old as (select * from CD009);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�G�٭�]�N�X��(CD011)���s�¥N�X��� Table: CD011_Map
DROP TABLE CD011_Map CASCADE CONSTRAINT;
create table CD011_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD011 �¥N�X Table *****
drop table CD011_Old cascade constraint;
create table CD011_Old as (select * from CD011);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�G�٭�]�N�X��2(CD011B)���s�¥N�X��� Table: CD011B_Map
DROP TABLE CD011B_Map CASCADE CONSTRAINT;
create table CD011B_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD011B �¥N�X Table *****
drop table CD011B_Old cascade constraint;
create table CD011B_Old as (select * from CD011B);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-���P��]�N�X��(CD012)���s�¥N�X��� Table: CD012_Map
DROP TABLE CD012_Map CASCADE CONSTRAINT;
create table CD012_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD012 �¥N�X Table *****
drop table CD012_Old cascade constraint;
create table CD012_Old as (select * from CD012);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-��ú�O��]�N�X��(CD013)���s�¥N�X��� Table: CD013_Map
DROP TABLE CD013_Map CASCADE CONSTRAINT;
create table CD013_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD013 �¥N�X Table *****
drop table CD013_Old cascade constraint;
create table CD013_Old as (select * from CD013);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�������]�N�X��(CD014)���s�¥N�X��� Table: CD014_Map
DROP TABLE CD014_Map CASCADE CONSTRAINT;
create table CD014_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD014 �¥N�X Table *****
drop table CD014_Old cascade constraint;
create table CD014_Old as (select * from CD014);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�h���]�N�X��(CD015)���s�¥N�X��� Table: CD015_Map
DROP TABLE CD015_Map CASCADE CONSTRAINT;
create table CD015_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD015 �¥N�X Table *****
drop table CD015_Old cascade constraint;
create table CD015_Old as (select * from CD015);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�u����]�N�X��(CD016)���s�¥N�X��� Table: CD016_Map
DROP TABLE CD016_Map CASCADE CONSTRAINT;
create table CD016_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD016 �¥N�X Table *****
drop table CD016_Old cascade constraint;
create table CD016_Old as (select * from CD016);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�Ȧ�N�X��(CD018)���s�¥N�X��� Table: CD018_Map
DROP TABLE CD018_Map CASCADE CONSTRAINT;
create table CD018_Map (
	NewCode number(3),
	NewName varchar2(24),
	OldCode number(3),
	OldName varchar2(24),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD018 �¥N�X Table *****
drop table CD018_Old cascade constraint;
create table CD018_Old as (select * from CD018);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-���O���إN�X��(CD019)���s�¥N�X��� Table: CD019_Map
DROP TABLE CD019_Map CASCADE CONSTRAINT;
create table CD019_Map (
	NewCode number(5),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD019 �¥N�X Table *****
drop table CD019_Old cascade constraint;
create table CD019_Old as (select * from CD019);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�I�O�N�@�N�X��(CD020)���s�¥N�X��� Table: CD020_Map
DROP TABLE CD020_Map CASCADE CONSTRAINT;
create table CD020_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD020 �¥N�X Table *****
drop table CD020_Old cascade constraint;
create table CD020_Old as (select * from CD020);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�ؿv���A�N�X��(CD021)���s�¥N�X��� Table: CD021_Map
DROP TABLE CD021_Map CASCADE CONSTRAINT;
create table CD021_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD021 �¥N�X Table *****
drop table CD021_Old cascade constraint;
create table CD021_Old as (select * from CD021);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�~�W�s���N�X��(CD022)���s�¥N�X��� Table: CD022_Map
DROP TABLE CD022_Map CASCADE CONSTRAINT;
create table CD022_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD022 �¥N�X Table *****
drop table CD022_Old cascade constraint;
create table CD022_Old as (select * from CD022);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�Ȥ�ӷ��N�X��(CD025)���s�¥N�X��� Table: CD025_Map
DROP TABLE CD025_Map CASCADE CONSTRAINT;
create table CD025_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD025 �¥N�X Table *****
drop table CD025_Old cascade constraint;
create table CD025_Old as (select * from CD025);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�A�Ⱥ��N�ץN�X��(CD026)���s�¥N�X��� Table: CD026_Map
DROP TABLE CD026_Map CASCADE CONSTRAINT;
create table CD026_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD026 �¥N�X Table *****
drop table CD026_Old cascade constraint;
create table CD026_Old as (select * from CD026);



--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�����ӷ���(CD049)���s�¥N�X��� Table: CD049_Map
DROP TABLE CD049_Map CASCADE CONSTRAINT;
create table CD049_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD049 �¥N�X Table *****
drop table CD049_Old cascade constraint;
create table CD049_Old as (select * from CD049);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�@�o��]�N�X��(CD051)���s�¥N�X��� Table: CD051_Map
DROP TABLE CD051_Map CASCADE CONSTRAINT;
create table CD051_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD051 �¥N�X Table *****
drop table CD051_Old cascade constraint;
create table CD051_Old as (select * from CD051);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�޽u���O�N�X��(CD055)���s�¥N�X��� Table: CD055_Map
DROP TABLE CD055_Map CASCADE CONSTRAINT;
create table CD055_Map (
	NewCode number(3),
	NewName varchar2(26),
	OldCode number(3),
	OldName varchar2(26),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD055 �¥N�X Table *****
drop table CD055_Old cascade constraint;
create table CD055_Old as (select * from CD055);


--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-���ث~�N�X��(CD059)���s�¥N�X��� Table: CD059_Map
DROP TABLE CD059_Map CASCADE CONSTRAINT;
create table CD059_Map (
	NewCode number(8),
	NewName varchar2(26),
	OldCode number(8),
	OldName varchar2(26),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD059 �¥N�X Table *****
drop table CD059_Old cascade constraint;
create table CD059_Old as (select * from CD059);





Prompt OK
exit




