/*
���ɫe������ Table��Index allocate space �[�����ɮɶ�
@C:\Transform\EMC_Transform\Index_Trigger\allocate.sql
*/


Prompt ���ɫe������ Table allocate space �[�����ɮɶ�
alter table So001 allocate extent (size 150M);
alter table So002 allocate extent (size 150M);
alter table So007 allocate extent (size 150M);
alter table So009 allocate extent (size 150M);
alter table So014 allocate extent (size 150M);
alter table So033 allocate extent (size 150M);
alter table So034 allocate extent (size 150M);


Prompt ���ɫe������ index allocate space �[�����ɮɶ�
alter index PK_So001 allocate extent (size 100M);
alter index PK_So002 allocate extent (size 100M);
alter index PK_So007 allocate extent (size 100M);
alter index PK_So009 allocate extent (size 100M);
alter index PK_So014 allocate extent (size 100M);
alter index PK_So033 allocate extent (size 100M);

alter index PK_So034 allocate extent (size 100M);