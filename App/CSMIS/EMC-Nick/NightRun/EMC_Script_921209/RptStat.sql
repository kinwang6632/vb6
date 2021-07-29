★so054(報表使用記錄資料檔)

(1)將 GetSec.sql compiler到後端 (用來換算so054.second欄位)。
(2)用FTool執行下列sql指令，並將結果存成檔案即可。

-- 編號, 名稱, 最大秒數, 最小秒數, 執次數
select prgname, so029.prompt prompt, max(GetSec(second)), min(GetSec(second)), count(*)
  from so054 , so029
  where substr(prgname,1,6)=so029.mid
  group by prgName, so029.prompt;


-- 編號, 名稱, 最大秒數, 條件
select a.PrgName, b.Prompt, a.second, c.selection
  from (select prgName,max(GetSec(second)) second from SO054 group by prgname) a, 
  so029 b, so054  c
  where substr(a.prgname,1,6)=b.mid and a.prgname=c.prgname and GetSec(c.second)=a.second
  order by a.prgname;


-- 編號, 名稱, 最小秒數, 條件
select a.PrgName, b.Prompt, a.second, c.selection
  from (select prgName,min(GetSec(second)) second from SO054 group by prgname) a, 
  so029 b, so054  c
  where substr(a.prgname,1,6)=b.mid and a.prgname=c.prgname and GetSec(c.second)=a.second
  order by a.prgname;


