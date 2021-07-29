-- *********************************************************************************
-- 建立EMC營運管理報表資料檔的script				Date: 2003.09.04
--   @c:\emc_script\GenSO511
--   註: 為了趕緊上線, 先不建立索引, 請人工事後補建立
-- *********************************************************************************
  prompt *** 建立EMC營運管理報表資料檔 ***

  prompt SO5D01/04/07營運管理報表資料檔: SO511A
  drop table SO511A;
  create table SO511A (
	CompCode number(3), RptYM number(6), ClassCode number(3),
	PeriodCode number(3), NextYM number(6), CustCount number(8), DataType number(1));

  prompt *** SO5D02收費單數預估統計表: SO511B
  drop table SO511B;
  create table SO511B (
	CompCode number(3), RptYM number(6), ClassCode number(3),
	PeriodCode number(3), NextYM number(6), CustCount number(8));

  prompt *** SO5D03收費金額預估統計表: SO511C
  drop table SO511C;
  create table SO511C (
	CompCode number(3), RptYM number(6), ClassCode number(3),
	PeriodCode number(3), NextYM number(6), Amount number(10));

  prompt 各繳費期別牌價表: SO511Z
  drop table SO511Z;
  create table SO511Z (
	CompCode number(3), ClassCode number(3), PeriodCode number(3), UnitPrice number(8));

  ----------------------------------------------------------------------
  prompt *** 建立牌價表內容: 請確定此檔案內容, 每一家不見得一樣 ***
  -- 以下牌價表係根據92.09.04 EMC李宗富提供之'系統台收費牌價表.xls'而定的
  ----------------------------------------------------------------------
  -- (1) 先取目前公司別
  variable v_CompCode number;
  exec :v_CompCode := 0;
  declare
  begin
    select CompCode into :v_CompCode from SO501A where RowNum<=1;
  end;
/
  prompt 目前公司別
  print v_CompCode;

  -- (2) 新增牌價表
  --	 1. 公關
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 1, 1, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 1, 2, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 1, 3, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 1, 4, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 1, 5, 0);

  --	 2. 供電
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 2, 1, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 2, 2, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 2, 3, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 2, 4, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 2, 5, 0);

  --	 3. 大樓公關或團體
  begin
    if :v_CompCode in (9,10,11,12) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 1, 275);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 2, 550);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 3, 820);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 4, 1650);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 5, 3250);
    elsif :v_CompCode = 8 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 1, 280);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 2, 560);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 3, 825);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 4, 1650);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 5, 3250);
    elsif :v_CompCode in (13,14) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 1, 280);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 2, 560);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 3, 840);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 4, 1650);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 5, 3250);
    elsif :v_CompCode = 7 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 1, 295);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 2, 590);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 3, 885);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 4, 1770);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 5, 3540);
    elsif :v_CompCode = 6 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 1, 290);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 2, 580);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 3, 870);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 4, 1740);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 5, 3480);
    elsif :v_CompCode = 5 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 1, 300);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 2, 600);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 3, 900);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 4, 1800);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 5, 3600);
    elsif :v_CompCode = 3 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 1, 300);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 2, 600);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 3, 900);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 4, 1800);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 5, 3500);
    elsif :v_CompCode in (1,2) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 1, 275);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 2, 550);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 3, 825);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 4, 1650);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 3, 5, 3300);
    end if;
  end;
/

  --	 4. 大樓統收,代表,公播
  begin
    if :v_CompCode in (9,10,11,12) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 1, 500);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 2, 1000);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 3, 1500);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 4, 3000);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 5, 6000);
    elsif :v_CompCode = 8 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 1, 400);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 2, 800);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 3, 1200);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 4, 2400);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 5, 4800);
    elsif :v_CompCode in (13,14) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 1, 500);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 2, 1000);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 3, 1500);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 4, 3000);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 5, 6000);
    elsif :v_CompCode = 7 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 1, 530);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 2, 1060);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 3, 1590);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 4, 3180);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 5, 6360);
    elsif :v_CompCode = 6 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 1, 520);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 2, 1040);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 3, 1560);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 4, 3120);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 5, 6240);
    elsif :v_CompCode = 5 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 1, 520);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 2, 1040);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 3, 1560);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 4, 3120);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 5, 6240);
    elsif :v_CompCode = 3 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 1, 583);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 2, 1167);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 3, 1750);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 4, 3500);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 5, 7000);
    elsif :v_CompCode in (1,2) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 1, 550);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 2, 1100);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 3, 1650);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 4, 3300);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 4, 5, 6600);
    end if;
  end;
/

  --	 5. 大樓個收
  begin
    if :v_CompCode in (9,10,11,12) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 1, 550);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 2, 1100);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 3, 1640);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 4, 3200);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 5, 6300);
    elsif :v_CompCode = 8 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 1, 400);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 2, 800);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 3, 1350);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 4, 2700);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 5, 4800);
    elsif :v_CompCode in (13,14) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 1, 560);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 2, 1120);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 3, 1680);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 4, 3300);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 5, 6100);
    elsif :v_CompCode = 7 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 1, 530);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 2, 1060);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 3, 1590);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 4, 3180);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 5, 6360);
    elsif :v_CompCode = 6 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 1, 520);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 2, 1040);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 3, 1560);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 4, 3120);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 5, 6240);
    elsif :v_CompCode = 5 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 1, 520);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 2, 1040);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 3, 1560);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 4, 3120);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 5, 6240);
    elsif :v_CompCode = 3 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 1, 600);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 2, 1200);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 3, 1800);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 4, 3600);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 5, 7000);
    elsif :v_CompCode in (1,2) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 1, 550);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 2, 1100);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 3, 1650);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 4, 3300);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 5, 5, 6600);
    end if;
  end;
/

  --	 6. 一般戶
  begin
    if :v_CompCode in (9,10,11,12) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 1, 550);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 2, 1100);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 3, 1640);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 4, 3300);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 5, 6500);
    elsif :v_CompCode = 8 then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 1, 560);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 2, 1120);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 3, 1650);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 4, 3300);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 5, 6500);
    elsif :v_CompCode in (13,14) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 1, 560);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 2, 1120);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 3, 1680);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 4, 3300);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 5, 6500);
    elsif :v_CompCode = 7 then
 	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 1, 590);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 2, 1180);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 3, 1770);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 4, 3540);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 5, 7080);
    elsif :v_CompCode = 6 then
 	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 1, 580);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 2, 1160);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 3, 1740);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 4, 3480);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 5, 6960);
    elsif :v_CompCode = 5 then
 	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 1, 600);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 2, 1200);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 3, 1800);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 4, 3600);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 5, 7200);
    elsif :v_CompCode = 3 then
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 1, 600);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 2, 1200);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 3, 1800);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 4, 3600);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 5, 7000);
    elsif :v_CompCode in (1,2) then
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 1, 550);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 2, 1100);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 3, 1650);
	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 4, 3300);
  	insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 6, 5, 6600);
    end if;
  end;
/

  --	 7. 其他
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 7, 1, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 7, 2, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 7, 3, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 7, 4, 0);
  insert into SO511Z (CompCode, ClassCode, PeriodCode, UnitPrice) values (:v_CompCode, 7, 5, 0);
  commit;

  prompt OK
