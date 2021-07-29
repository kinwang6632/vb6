/*
    補充更多的數位電視上線所需代碼, 切至各SO資料區直接執行此檔案即可
    Date: 2003.01.16
    By: Pierre
*/

-- 取公司別
variable v_CompCode number
begin
    select CodeNo into :v_CompCode from CD039 where rownum<=1;
end;
/

-----------------------------------------------------------------------
prompt CD061 STB結清原因代碼
delete from CD061;
insert into CD061 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 1, '搬家不看了', 0);
insert into CD061 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 2, '節目問題', 0);
insert into CD061 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 3, '不想看', 0);
insert into CD061 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 4, '其他', 0);


-----------------------------------------------------------------------
prompt CD060 頻道結清原因代碼
delete from CD060;
insert into CD060 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 1, '節目問題', 0);
insert into CD060 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 2, '費用問題', 0);
insert into CD060 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 3, '不想看', 0);
insert into CD060 (CompCode, CodeNo, Description, StopFlag) 
    values (:v_CompCode, 4, '其他', 0);


-----------------------------------------------------------------------
prompt CD028 STB配件代碼檔
delete from CD028;
insert into CD028 (CompCode, CodeNo, Description, Cost, StopFlag)
    values (:v_CompCode, 1, '智慧卡', 800, 0);
insert into CD028 (CompCode, CodeNo, Description, Cost, StopFlag)
    values (:v_CompCode, 2, '遙控器', 500, 0);
insert into CD028 (CompCode, CodeNo, Description, Cost, StopFlag)
    values (:v_CompCode, 3, '電源線', 200, 0);
insert into CD028 (CompCode, CodeNo, Description, Cost, StopFlag)
    values (:v_CompCode, 4, 'AV線', 200, 0);


-----------------------------------------------------------------------
prompt CD057 訂單類別代碼檔
delete from CD057;
insert into CD057 (CodeNo, Description, RefNo, StopFlag)
    values (1, '申裝eBox', 10, 0);
insert into CD057 (CodeNo, Description, RefNo, StopFlag)
    values (2, '付費頻道異動', 9, 0);


-----------------------------------------------------------------------
prompt CD051 作廢原因代碼檔
declare
  v_Cnt number := 0;
  v_MaxCodeNo number := 0;
begin
  select count(*) into v_Cnt from CD051 where RefNo=1;
  select max(nvl(CodeNo,0)) into v_MaxCodeNo from CD051;

  if v_Cnt =0 then
    insert into CD051 (CompCode, CodeNo, Description, RefNo, StopFlag)
	values (:v_CompCode, v_MaxCodeNo+1, '頻道結清', 1, 0);
  end if;
end;
/

