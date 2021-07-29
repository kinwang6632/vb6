spool c:\temp\RemoveMvLog.log

prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;

prompt   注意: 請將 EMCUC 修改成適當的 owner
prompt   注意: 請修改 system 的密碼以及連線的 alias

CONNECT system/manager@tc;

prompt CD001 ~ CD010


DROP MATERIALIZED VIEW LOG ON EMCTC.CD000 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD001 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD001A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD002 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD002CM003 ;

DROP MATERIALIZED VIEW LOG ON EMCTC.CD003 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD004 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD005 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD006 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD006A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD006B ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD007 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD008 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD008A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD008B ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD008C ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD009 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD010 ;


prompt CD011 ~ Cd020

DROP MATERIALIZED VIEW LOG ON EMCTC.CD011 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD011A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD011B ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD012 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD013 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD014 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD015 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD016 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD017 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD018 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD019 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD019A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD019CD004 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD019SO017 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD020 ;


prompt CD021 ~ Cd030


DROP MATERIALIZED VIEW LOG ON EMCTC.CD021 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD022 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD022CD019 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD023 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD024 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD025 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD026 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD027 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD027A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD027B ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD027C ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD027D ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD028 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD029 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD030 ;

prompt CD031 ~ Cd040


DROP MATERIALIZED VIEW LOG ON EMCTC.CD031 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD031A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD032 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD033 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD034 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD035 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD036 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD037 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD039 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD040 ;



prompt CD041 ~ Cd050


DROP MATERIALIZED VIEW LOG ON EMCTC.CD041 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD042 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD043 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD044 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD045 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD046 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD047 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD048 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD048A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD049 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD050 ;
  		

prompt Cd051 以後


DROP MATERIALIZED VIEW LOG ON EMCTC.CD051 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD052 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD053 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD054 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD055 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD056 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD057 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD058 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD059 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD060 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CD061 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CM001 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CM002 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CM003 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CM004 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.CM005 ;



prompt So001~So010


DROP MATERIALIZED VIEW LOG ON EMCTC.SO001 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO002 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO002A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO003 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO004 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO005 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO006 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO007 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO008 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO009 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO010 ;




prompt So011~So020


DROP MATERIALIZED VIEW LOG ON EMCTC.SO011 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO012 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO013 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO014 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO015 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO016 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO017A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO019 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO020 ;





prompt So021~So030


DROP MATERIALIZED VIEW LOG ON EMCTC.SO021 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO022 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO023 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO024 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO025 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO026 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO027 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO028 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO029 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO029A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO029T ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO030 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO030A ;




prompt  So031~So040


DROP MATERIALIZED VIEW LOG ON EMCTC.SO031 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO032 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO032BREAK ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO033 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO034 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO035 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO036 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO037 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO039 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO040 ;




prompt So041~So080


DROP MATERIALIZED VIEW LOG ON EMCTC.SO041 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO042 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO043 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO044 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO045 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO046 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO052 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO054 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO062 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO063 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO064 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO065 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO066 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO067 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO069 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO074 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO075 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO077 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO079 ;



prompt So081~So0100


DROP MATERIALIZED VIEW LOG ON EMCTC.SO081 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO082 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO083 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO084 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO085 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO086 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO087 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO088 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO089 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO090 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO091 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO092 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO093 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO094 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO096 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO098 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO100 ;




prompt So101~So130

DROP MATERIALIZED VIEW LOG ON EMCTC.SO102 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO103 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO104 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO105 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO105A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO105B ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO105C ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO105D ;

DROP MATERIALIZED VIEW LOG ON EMCTC.SO106 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO107 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO110 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO111 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO112 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO113 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO114 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO115 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO116 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO117 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO126 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO127 ;





prompt So200~So600

DROP MATERIALIZED VIEW LOG ON EMCTC.SO200 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO201 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO301A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO301B ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO302A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO302B ;

DROP MATERIALIZED VIEW LOG ON EMCTC.SO303 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO304 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO501 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO501A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO501X ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO502 with  rowid;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO502X ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SO503 ;


prompt others

DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0101 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0102 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0103 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0201A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0201B ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0201C ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0201D ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0201E ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0202 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0203A ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0203B ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOAC0204 ;
DROP MATERIALIZED VIEW LOG ON EMCTC.SOPOWER ;

prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;

spool off
