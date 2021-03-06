spool c:\temp\CreateMvLog.log

prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;

prompt   注意: 請將 EMCUC 修改成適當的 owner
prompt   注意: 請修改 system 的密碼以及連線的 alias
prompt   注意: 若該 table 沒有  primary key, 則 with primary key 必須改為with rowid.
prompt   注意: 若要取消此項設定,範例指令如下:SQL>DROP MATERIALIZED VIEW LOG ON EMCKP.CD008;

CONNECT system/manager@tc;

prompt CD001 ~ CD010


 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD000 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD001 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD001A with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD002 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD002CM003 with primary key;

 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD003 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD004 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD005 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD006 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD006A with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD006B with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD007 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD008 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD008A with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD008B with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD008C with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD009 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD010 with primary key;


prompt CD011 ~ Cd020

 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD011 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD011A with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD011B with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD012 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD013 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD014 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD015 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD016 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD017 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD018 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD019 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD019A with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD019CD004 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD019SO017 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD020 with primary key;


prompt CD021 ~ Cd030


 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD021 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD022 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD022CD019 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD023 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD024 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD025 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD026 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD027 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD027A with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD027B with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD027C with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD027D with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD028 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD029 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD030 with primary key;

prompt CD031 ~ Cd040


 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD031 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD031A with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD032 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD033 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD034 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD035 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD036 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD037 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD039 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD040 with primary key;



prompt CD041 ~ Cd050


 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD041 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD042 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD043 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD044 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD045 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD046 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD047 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD048 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD048A with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD049 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD050 with primary key;
  		

prompt Cd051 以後


 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD051 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD052 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD053 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD054 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD055 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD056 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD057 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD058 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD059 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD060 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD061 with primary key;

 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD062 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CD063 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CM001 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CM002 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CM003 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CM004 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.CM005 with rowid;



prompt So001~So010


 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO001 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO002 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO002A with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO003 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO004 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO005 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO006 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO007 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO008 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO009 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO010 with primary key;




prompt So011~So020


 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO011 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO012 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO013 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO014 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO015 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO016 with primary key;

 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO017 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO017A with rowid;

 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO018 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO019 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO020 with primary key;





prompt So021~So030


 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO021 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO022 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO023 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO024 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO025 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO026 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO027 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO028 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO029 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO029A with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO030 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO030A with rowid;




prompt  So031~So040


 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO031 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO032 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO032BREAK with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO033 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO034 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO035 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO036 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO037 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO038 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO039 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO040 with rowid;




prompt So041~So080


 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO041 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO042 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO043 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO044 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO045 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO046 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO052 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO054 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO062 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO063 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO064 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO065 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO066 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO067 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO069 with primary key;

 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO071 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO072 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO073 with primary key;


 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO074 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO075 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO077 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO079 with rowid;



prompt So081~So0100


 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO081 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO082 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO083 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO084 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO085 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO086 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO087 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO088 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO089 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO090 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO091 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO092 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO093 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO094 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO095 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO096 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO097 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO098 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO099 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO100 with rowid;







prompt So101~So130

 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO102 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO103 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO104 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO105 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO105A with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO105B with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO105C with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO105D with rowid;

 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO106 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO107 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO108 with rowid;

 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO110 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO111 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO112 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO113 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO114 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO115 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO116 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO117 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO118 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO119 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO123 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO124 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO126 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO127 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO128 with rowid;

prompt So131~So140

 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO131 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO132 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO133 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO134 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO135 with primary key;


prompt So200~So600

 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO200 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO201 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO301A with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO301B with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO302A with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO302B with rowid;

 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO303 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO304 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO501 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO501A with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO501X with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO502 with  rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO502X with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SO503 with rowid;


prompt others

 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0101 with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0102 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0103 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0201A with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0201B with primary key;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0201C with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0201D with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0201E with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0202 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0203A with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0203B with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOAC0204 with rowid;
 CREATE MATERIALIZED VIEW LOG ON EMCKP.SOPOWER with rowid;

prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;

spool off
