��so054(����ϥΰO�������)

(1)�N GetSec.sql compiler���� (�ΨӴ���so054.second���)�C
(2)��FTool����U�Csql���O�A�ñN���G�s���ɮקY�i�C

-- �s��, �W��, �̤j���, �̤p���, ������
select prgname, so029.prompt prompt, max(GetSec(second)), min(GetSec(second)), count(*)
  from so054 , so029
  where substr(prgname,1,6)=so029.mid
  group by prgName, so029.prompt;


-- �s��, �W��, �̤j���, ����
select a.PrgName, b.Prompt, a.second, c.selection
  from (select prgName,max(GetSec(second)) second from SO054 group by prgname) a, 
  so029 b, so054  c
  where substr(a.prgname,1,6)=b.mid and a.prgname=c.prgname and GetSec(c.second)=a.second
  order by a.prgname;


-- �s��, �W��, �̤p���, ����
select a.PrgName, b.Prompt, a.second, c.selection
  from (select prgName,min(GetSec(second)) second from SO054 group by prgname) a, 
  so029 b, so054  c
  where substr(a.prgname,1,6)=b.mid and a.prgname=c.prgname and GetSec(c.second)=a.second
  order by a.prgname;


