Prompt ***** 91.11.21 �إ߷s�¥N�X��� Table , �s�N�X Table �� �ƥ��¥N�X�� *****
--*******************************************************************
--*******************************************************************
Prompt �إ߫Ȥ����O�N�X��(CD004)���s�¥N�X��� Table: CD004_Map
DROP TABLE CD004_Map CASCADE CONSTRAINT;
create table CD004_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt �إ߫Ȥ����O�N�X��(CD004)���s�N�X Table: CD004_New
DROP TABLE CD004_New CASCADE CONSTRAINT;
create table CD004_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** �ƥ� CD004 �¥N�X Table *****
drop table CD004_Old cascade constraint;
create table CD004_Old as (select * from CD004);

--*******************************************************************
--*******************************************************************
Prompt �إߦ��O���إN�X��(CD019)���s�¥N�X��� Table: CD019_Map
DROP TABLE CD019_Map CASCADE CONSTRAINT;
create table CD019_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt �إ߫Ȥ����O�N�X��(CD019)���s�N�X Table: CD019_New
DROP TABLE CD019_New CASCADE CONSTRAINT;
create table CD019_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** �ƥ� CD019 �¥N�X  Table *****
drop table CD019_Old cascade constraint;
create table CD019_Old as (select * from CD019);

--*******************************************************************
--*******************************************************************
Prompt �إ߸˾����O�N�X��(CD005)���s�¥N�X��� Table: CD005_Map
DROP TABLE CD005_Map CASCADE CONSTRAINT;
create table CD005_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt �إ߫Ȥ����O�N�X��(CD005)���s�N�X Table: CD005_New
DROP TABLE CD005_New CASCADE CONSTRAINT;
create table CD005_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** �ƥ� CD005 �¥N�X  Table *****
drop table CD005_Old cascade constraint;
create table CD005_Old as (select * from CD005);

--*******************************************************************
--*******************************************************************
/*
Prompt �إߦ��O�覡�N�X��(CD031)���s�¥N�X��� Table: CD031_Map
DROP TABLE CD031_Map CASCADE CONSTRAINT;
create table CD031_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt �إ߫Ȥ����O�N�X��(CD031)���s�N�X Table: CD031_New
DROP TABLE CD031_New CASCADE CONSTRAINT;
create table CD031_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** �ƥ� CD031 �¥N�X  Table *****
drop table CD031_Old cascade constraint;
create table CD031_Old as (select * from CD031);
*/

--*******************************************************************
--*******************************************************************
/*
Prompt �إ߶R��覡�N�X��(CD034)���s�¥N�X��� Table: CD034_Map
DROP TABLE CD034_Map CASCADE CONSTRAINT;
create table CD034_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt �إ߫Ȥ����O�N�X��(CD034)���s�N�X Table: CD034_New
DROP TABLE CD034_New CASCADE CONSTRAINT;
create table CD034_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** �ƥ� CD034 �¥N�X  Table *****
drop table CD034_Old cascade constraint;
create table CD034_Old as (select * from CD034);
*/
--*******************************************************************
--*******************************************************************
Prompt �إ߰�������O�N�X��(CD007)���s�¥N�X���  Table: CD007_Map
DROP TABLE CD007_Map CASCADE CONSTRAINT;
create table CD007_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt �إ߫Ȥ����O�N�X��(CD007)���s�N�X Table: CD007_New
DROP TABLE CD007_New CASCADE CONSTRAINT;
create table CD007_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** �ƥ� CD007 �¥N�X  Table *****
drop table CD007_Old cascade constraint;
create table CD007_Old as (select * from CD007);

--*******************************************************************
--*******************************************************************
Prompt �إߥI�O�N�@�N�X��(CD020)���s�¥N�X��� Table: CD020_Map
DROP TABLE CD020_Map CASCADE CONSTRAINT;
create table CD020_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt �إ߫Ȥ����O�N�X��(CD020)���s�N�X Table: CD020_New
DROP TABLE CD020_New CASCADE CONSTRAINT;
create table CD020_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** �ƥ� CD020 �¥N�X  Table *****
drop table CD020_Old cascade constraint;
create table CD020_Old as (select * from CD020);

--*******************************************************************
--*******************************************************************
Prompt �إ߫ؿv���A�N�X��(CD021)���s�¥N�X��� Table: CD021_Map
DROP TABLE CD021_Map CASCADE CONSTRAINT;
create table CD021_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt �إ߫Ȥ����O�N�X��(CD021)���s�N�X Table: CD021_New
DROP TABLE CD021_New CASCADE CONSTRAINT;
create table CD021_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** �ƥ� CD021 �¥N�X  Table *****
drop table CD021_Old cascade constraint;
create table CD021_Old as (select * from CD021);

--*******************************************************************
--*******************************************************************
Prompt �إߴC�餶�����O�N�X��(CD009)���s�¥N�X��� Table: CD009_Map
DROP TABLE CD009_Map CASCADE CONSTRAINT;
create table CD009_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt �إ߫Ȥ����O�N�X��(CD009)���s�N�X Table: CD009_New
DROP TABLE CD009_New CASCADE CONSTRAINT;
create table CD009_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** �ƥ� CD009 �¥N�X  Table *****
drop table CD009_Old cascade constraint;
create table CD009_Old as (select * from CD009);



Prompt OK
exit




