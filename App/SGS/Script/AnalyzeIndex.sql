/*
Analyze V2.x Table and Index
Log file: C:\profiles\cablesoft\Tools\AnalyzeIndex.log
���O1: analyze table
             analyze table ??? estimate statistics;
���O2: analyze index
             analyze index ??? estimate statistics;
���O3: check analyze status
             select table_name, last_analyzed  from user_tables order by Table_name;

  By:  Jackal
  Date: 2004.04.06
*/




variable nn number
variable msg varchar2(200)
variable t1 varchar2(19)
variable t2 varchar2(19)
variable str varchar2(80)
variable tt number
variable hh number
variable mm number
variable ss number
set serveroutput on
set heading off
exec :t1 := to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS')
exec :str := '<< �}�l Analyze >> ' || :t1
spool C:\AnalyzeIndex.log
print str




prompt ���b�B�z I_SO004_SSBFP...
analyze index I_SO004_SSBFP compute statistics;

prompt ���b�B�z I_SO004_CustId_FaciCode...
analyze index I_SO004_CustId_FaciCode compute statistics;

prompt ���b�B�z I_SO004_FaciSNo...
analyze index I_SO004_FaciSNo compute statistics;

prompt ���b�B�z I_SO004_CSS...
analyze index I_SO004_CSS compute statistics;

prompt ���b�B�z I_SO004_FaciCode...
analyze index I_SO004_FaciCode compute statistics;

prompt ���b�B�z I_SO004_BuyCode...
analyze index I_SO004_BuyCode compute statistics;

prompt ���b�B�z I_SO004_InstDate...
analyze index I_SO004_InstDate compute statistics;

prompt ���b�B�z I_SO451_HandleTime...
analyze index I_SO451_HandleTime compute statistics;

prompt ���b�B�z I_SO453_HandleTime...
analyze index I_SO453_HandleTime compute statistics;

prompt ���b�B�z I_SO455_HandleTime...
analyze index I_SO455_HandleTime compute statistics;

prompt ���b�B�z I_SO457_HandleTime...
analyze index I_SO457_HandleTime compute statistics;




exec :t2 := to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
exec :str := '<< Analyze���槹��, ���ˬdlog��, �ݬO�_�����~ >> ' || :t2;
print str
exec :tt := TRUNC(86400*(to_date(:t2, 'YYYY/MM/DD HH24:MI:SS')-to_date(:t1, 'YYYY/MM/DD HH24:MI:SS')));
exec :hh := trunc(:tt/3600);
exec :ss := mod(:tt, 60);
exec :mm := (:tt - :hh*3600 - :ss) / 60;
exec :str := '�X�p����ɶ�: ' || :hh || '��' || :mm || '��' || :ss || '��';
print str
spool off
set heading on
