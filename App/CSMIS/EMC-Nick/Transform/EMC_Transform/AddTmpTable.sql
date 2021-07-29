Prompt ***** 91.11.21 建立新舊代碼對照 Table , 新代碼 Table 及 備份舊代碼檔 *****
--*******************************************************************
--*******************************************************************
Prompt 建立客戶類別代碼檔(CD004)之新舊代碼對照 Table: CD004_Map
DROP TABLE CD004_Map CASCADE CONSTRAINT;
create table CD004_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt 建立客戶類別代碼檔(CD004)之新代碼 Table: CD004_New
DROP TABLE CD004_New CASCADE CONSTRAINT;
create table CD004_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** 備份 CD004 舊代碼 Table *****
drop table CD004_Old cascade constraint;
create table CD004_Old as (select * from CD004);

--*******************************************************************
--*******************************************************************
Prompt 建立收費項目代碼檔(CD019)之新舊代碼對照 Table: CD019_Map
DROP TABLE CD019_Map CASCADE CONSTRAINT;
create table CD019_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt 建立客戶類別代碼檔(CD019)之新代碼 Table: CD019_New
DROP TABLE CD019_New CASCADE CONSTRAINT;
create table CD019_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** 備份 CD019 舊代碼  Table *****
drop table CD019_Old cascade constraint;
create table CD019_Old as (select * from CD019);

--*******************************************************************
--*******************************************************************
Prompt 建立裝機類別代碼檔(CD005)之新舊代碼對照 Table: CD005_Map
DROP TABLE CD005_Map CASCADE CONSTRAINT;
create table CD005_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt 建立客戶類別代碼檔(CD005)之新代碼 Table: CD005_New
DROP TABLE CD005_New CASCADE CONSTRAINT;
create table CD005_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** 備份 CD005 舊代碼  Table *****
drop table CD005_Old cascade constraint;
create table CD005_Old as (select * from CD005);

--*******************************************************************
--*******************************************************************
/*
Prompt 建立收費方式代碼檔(CD031)之新舊代碼對照 Table: CD031_Map
DROP TABLE CD031_Map CASCADE CONSTRAINT;
create table CD031_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt 建立客戶類別代碼檔(CD031)之新代碼 Table: CD031_New
DROP TABLE CD031_New CASCADE CONSTRAINT;
create table CD031_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** 備份 CD031 舊代碼  Table *****
drop table CD031_Old cascade constraint;
create table CD031_Old as (select * from CD031);
*/

--*******************************************************************
--*******************************************************************
/*
Prompt 建立買賣方式代碼檔(CD034)之新舊代碼對照 Table: CD034_Map
DROP TABLE CD034_Map CASCADE CONSTRAINT;
create table CD034_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt 建立客戶類別代碼檔(CD034)之新代碼 Table: CD034_New
DROP TABLE CD034_New CASCADE CONSTRAINT;
create table CD034_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** 備份 CD034 舊代碼  Table *****
drop table CD034_Old cascade constraint;
create table CD034_Old as (select * from CD034);
*/
--*******************************************************************
--*******************************************************************
Prompt 建立停拆移機類別代碼檔(CD007)之新舊代碼對照  Table: CD007_Map
DROP TABLE CD007_Map CASCADE CONSTRAINT;
create table CD007_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt 建立客戶類別代碼檔(CD007)之新代碼 Table: CD007_New
DROP TABLE CD007_New CASCADE CONSTRAINT;
create table CD007_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** 備份 CD007 舊代碼  Table *****
drop table CD007_Old cascade constraint;
create table CD007_Old as (select * from CD007);

--*******************************************************************
--*******************************************************************
Prompt 建立付費意願代碼檔(CD020)之新舊代碼對照 Table: CD020_Map
DROP TABLE CD020_Map CASCADE CONSTRAINT;
create table CD020_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt 建立客戶類別代碼檔(CD020)之新代碼 Table: CD020_New
DROP TABLE CD020_New CASCADE CONSTRAINT;
create table CD020_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** 備份 CD020 舊代碼  Table *****
drop table CD020_Old cascade constraint;
create table CD020_Old as (select * from CD020);

--*******************************************************************
--*******************************************************************
Prompt 建立建築型態代碼檔(CD021)之新舊代碼對照 Table: CD021_Map
DROP TABLE CD021_Map CASCADE CONSTRAINT;
create table CD021_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt 建立客戶類別代碼檔(CD021)之新代碼 Table: CD021_New
DROP TABLE CD021_New CASCADE CONSTRAINT;
create table CD021_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** 備份 CD021 舊代碼  Table *****
drop table CD021_Old cascade constraint;
create table CD021_Old as (select * from CD021);

--*******************************************************************
--*******************************************************************
Prompt 建立媒體介紹類別代碼檔(CD009)之新舊代碼對照 Table: CD009_Map
DROP TABLE CD009_Map CASCADE CONSTRAINT;
create table CD009_Map (
	OldCode number(3),
	OldName varchar2(20),
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt 建立客戶類別代碼檔(CD009)之新代碼 Table: CD009_New
DROP TABLE CD009_New CASCADE CONSTRAINT;
create table CD009_New (
	NewCode number(3),
	NewName varchar2(20),
        Mark number(1));

Prompt ***** 備份 CD009 舊代碼  Table *****
drop table CD009_Old cascade constraint;
create table CD009_Old as (select * from CD009);



Prompt OK
exit




