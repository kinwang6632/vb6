/*
轉檔後執行 Table及Index deallocate 用以收回沒用到的 space
@C:\Transform\EMC_Transform\Index_Trigger\deallocate.sql
*/


Prompt 轉檔後執行 Table deallocate 用以收回沒用到的 space
alter table  so001 deallocate unused;
alter table  so002 deallocate unused;
alter table  so007 deallocate unused;
alter table  so009 deallocate unused;
alter table  so014 deallocate unused;
alter table  so033 deallocate unused;
alter table  so034 deallocate unused;

Prompt 轉檔前先執行 index allocate 用以收回沒用到的 space
alter index PK_So001 deallocate unused;
alter index PK_So002 deallocate unused;
alter index PK_So014 deallocate unused;
alter index PK_So007 deallocate unused;
alter index PK_So009 deallocate unused;
alter index PK_So033 deallocate unused;
alter index PK_So034 deallocate unused;