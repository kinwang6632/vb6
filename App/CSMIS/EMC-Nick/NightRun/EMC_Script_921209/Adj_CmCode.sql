/*
  �վ�EMC�U�a�����O�覡�N�X, �H�Φ��Ψ즬�O�覡�������. 

  �վ�覡:
  < �e�m�@�~ >
	1. ���X�C�@�a�����O�覡�N�X�s�­ȹ�Ӫ�
	2. �̾ڹ�Ӫ�, �g�X�Ӯa�����ഫscript file

  < �{���@�~ >
	1. login�C�@�a
	2. compile��N�X��ƨ禡
	3. ����Ӯa���N�X�ഫ�{��

  �`�N�ƶ�:
  . �]���i��վ�j�q���, �G�i���wrollback segment
*/

set serveroutput on
set heading off
variable str varchar2(80)
exec :str := '<< �վ�EMC�U�a�����O�覡�N�X >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool c:\gird\Log\Adj_CmCode.log
print str


-- 1. �վ�EMCYMS
conn EMCYMS/EMCYMS@????
@c:\gird\v300\script\SP_UpdCode
exec SP_UpdCode('xxx', a, b);
set transaction use rollback segment rbS27;


spool off
set heading on
