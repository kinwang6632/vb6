/*
���ɫ���� Table��Index deallocate �ΥH���^�S�Ψ쪺 space
@C:\Transform\EMC_Transform\Index_Trigger\deallocate.sql
*/


Prompt ���ɫ���� Table deallocate �ΥH���^�S�Ψ쪺 space
alter table  so001 deallocate unused;
alter table  so002 deallocate unused;
alter table  so007 deallocate unused;
alter table  so009 deallocate unused;
alter table  so014 deallocate unused;
alter table  so033 deallocate unused;
alter table  so034 deallocate unused;

Prompt ���ɫe������ index allocate �ΥH���^�S�Ψ쪺 space
alter index PK_So001 deallocate unused;
alter index PK_So002 deallocate unused;
alter index PK_So014 deallocate unused;
alter index PK_So007 deallocate unused;
alter index PK_So009 deallocate unused;
alter index PK_So033 deallocate unused;
alter index PK_So034 deallocate unused;