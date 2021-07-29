/*
Analyze V2.x Table and Index
Log file: C:\profiles\cablesoft\Tools\AnalyzeTable.log
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
spool C:\AnalyzeTable.log
print str



prompt ���b�B�z SO004...
analyze table SO004 compute statistics;

prompt ���b�B�z SOAC0201A...
analyze table SOAC0201A compute statistics;

prompt ���b�B�z SOAC0201B...
analyze table SOAC0201B compute statistics;

prompt ���b�B�z SO451...
analyze table SO451 compute statistics;

prompt ���b�B�z SO452...
analyze table SO452 compute statistics;

prompt ���b�B�z SO453...
analyze table SO453 compute statistics;

prompt ���b�B�z SO454...
analyze table SO454 compute statistics;

prompt ���b�B�z SO455...
analyze table SO455 compute statistics;

prompt ���b�B�z SO456...
analyze table SO456 compute statistics;

prompt ���b�B�z SO457...
analyze table SO457 compute statistics;

prompt ���b�B�z SO458...
analyze table SO458 compute statistics;

prompt ���b�B�z SO459...
analyze table SO459 compute statistics;

prompt ���b�B�z SO460...
analyze table SO460 compute statistics;

prompt ���b�B�z SO461...
analyze table SO461 compute statistics;



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
