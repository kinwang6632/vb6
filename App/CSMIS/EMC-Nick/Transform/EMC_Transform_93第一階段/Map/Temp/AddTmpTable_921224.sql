Prompt ***** 92.12.21 �إ߷s�¥N�X��� Table �� �ƥ��¥N�X�� *****
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


--*******************************************************************
--*******************************************************************
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
Prompt �إ�-���ץӧi���O���p��(CD006A)���s�¥N�X��� Table: CD006A_Map
DROP TABLE CD006A_Map CASCADE CONSTRAINT;
create table CD006A_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD006A �¥N�X Table *****
drop table CD006A_Old cascade constraint;
create table CD006A_Old as (select * from CD006A);

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
Prompt �إ�-�ӧi���O���ӧi���e(CD008C)���s�¥N�X��� Table: CD008C_Map
DROP TABLE CD008C_Map CASCADE CONSTRAINT;
create table CD008C_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD008C �¥N�X Table *****
drop table CD008C_Old cascade constraint;
create table CD008C_Old as (select * from CD008C);

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
Prompt �إ�-�M�\�N�X��(CD027)���s�¥N�X��� Table: CD027_Map
DROP TABLE CD027_Map CASCADE CONSTRAINT;
create table CD027_Map (
	NewCode varchar2(3),
	NewName varchar2(26),
	OldCode varchar2(3),
	OldName varchar2(26),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD027 �¥N�X Table *****
drop table CD027_Old cascade constraint;
create table CD027_Old as (select * from CD027);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-���O�覡�N�X��(CD031)���s�¥N�X��� Table: CD031_Map
DROP TABLE CD031_Map CASCADE CONSTRAINT;
create table CD031_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD031 �¥N�X Table *****
drop table CD031_Old cascade constraint;
create table CD031_Old as (select * from CD031);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�I�ں����N�X��(CD032)���s�¥N�X��� Table: CD032_Map
DROP TABLE CD032_Map CASCADE CONSTRAINT;
create table CD032_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD032 �¥N�X Table *****
drop table CD032_Old cascade constraint;
create table CD032_Old as (select * from CD032);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�R��覡�N�X��(CD034)���s�¥N�X��� Table: CD034_Map
DROP TABLE CD034_Map CASCADE CONSTRAINT;
create table CD034_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD034 �¥N�X Table *****
drop table CD034_Old cascade constraint;
create table CD034_Old as (select * from CD034);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�~�Ȭ��ʥN�X��(CD042)���s�¥N�X��� Table: CD042_Map
DROP TABLE CD042_Map CASCADE CONSTRAINT;
create table CD042_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD042 �¥N�X Table *****
drop table CD042_Old cascade constraint;
create table CD042_Old as (select * from CD042);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�ެ����O�N�X��(CD044)���s�¥N�X��� Table: CD044_Map
DROP TABLE CD044_Map CASCADE CONSTRAINT;
create table CD044_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD044 �¥N�X Table *****
drop table CD044_Old cascade constraint;
create table CD044_Old as (select * from CD044);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�u�ݿ�k�N�X��(CD048)���s�¥N�X��� Table: CD048_Map
DROP TABLE CD048_Map CASCADE CONSTRAINT;
create table CD048_Map (
	NewCode number(10),
	NewName varchar2(20),
	OldCode number(10),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD048 �¥N�X Table *****
drop table CD048_Old cascade constraint;
create table CD048_Old as (select * from CD048);

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
Prompt �إ�-�]�Ƹ˸m�I�N�X��(CD056)���s�¥N�X��� Table: CD056_Map
DROP TABLE CD056_Map CASCADE CONSTRAINT;
create table CD056_Map (
	NewCode number(3),
	NewName varchar2(26),
	OldCode number(3),
	OldName varchar2(26),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD056 �¥N�X Table *****
drop table CD056_Old cascade constraint;
create table CD056_Old as (select * from CD056);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�q�����O�N�X��(CD057)���s�¥N�X��� Table: CD057_Map
DROP TABLE CD057_Map CASCADE CONSTRAINT;
create table CD057_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD057 �¥N�X Table *****
drop table CD057_Old cascade constraint;
create table CD057_Old as (select * from CD057);

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

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-BISP�u�����ǥN�X��(CD062)���s�¥N�X��� Table: CD062_Map
DROP TABLE CD062_Map CASCADE CONSTRAINT;
create table CD062_Map (
	NewCode varchar2(20),
	NewName varchar2(20),
	OldCode varchar2(20),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD062 �¥N�X Table *****
drop table CD062_Old cascade constraint;
create table CD062_Old as (select * from CD062);

--%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Prompt �إ�-�K�X���ܥN�X��(CD063)���s�¥N�X��� Table: CD063_Map
DROP TABLE CD063_Map CASCADE CONSTRAINT;
create table CD063_Map (
	NewCode number(3),
	NewName varchar2(20),
	OldCode number(3),
	OldName varchar2(20),
	RefNo   number(3),
	Mark   number(3)
	);

Prompt ***** �ƥ� CD063 �¥N�X Table *****
drop table CD063_Old cascade constraint;
create table CD063_Old as (select * from CD063);







Prompt OK
exit




