Prompt ***** 92.12.27 �إ߷s�¥N�X��� Table �� �ƥ��¥N�X�� *****
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








Prompt OK
exit




