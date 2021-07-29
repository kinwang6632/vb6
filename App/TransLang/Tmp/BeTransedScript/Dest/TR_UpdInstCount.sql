 
-- 新增
create or replace trigger TR_SO001_InstCount1
  after insert on SO007
  for each row
  declare
  v_RefNo Number;
  begin
    select RefNo into v_RefNo from cd005 where CodeNo=:new.InstCode;
    v_RefNo := nvl(v_RefNo,0);
 
    -- 新增"完工資料且裝機台數!=0"才異動InstCount
    if :new.FinTime is not null and nvl(:new.InstCount,0)!=0 then
        if v_RefNo=1 or v_RefNo=5 then
           update SO001 set InstCount = nvl(:new.InstCount,0)
             where CustId = :new.CustId; 
        else
           update SO001 set InstCount = nvl(InstCount,0) + nvl(:new.InstCount,0)
             where CustId = :new.CustId; 
        end if;
    end if;
  end;
/
 
-- 修改
create or replace trigger TR_SO001_InstCount2
  after update on SO007
  for each row
  declare
  v_RefNo Number;
  begin
    select RefNo into v_RefNo from cd005 where CodeNo=:new.InstCode;
    v_RefNo := nvl(v_RefNo,0);
 
    -- 更新"完工資料且裝機台數!=0"才異動InstCount
    if :old.FinTime is null and :new.FinTime is not null and nvl(:new.InstCount,0)!=0 then
        if v_RefNo=1 or v_RefNo=5 then
            update SO001 set InstCount = nvl(:new.InstCount,0)
                where CustId = :new.CustId;
        else
            update SO001 set InstCount = nvl(InstCount,0) + nvl(:new.InstCount,0)
                where CustId = :new.CustId;
        end if;
    -- 更新"完工資料且裝機台數不同"才異動InstCount
    elsif :old.FinTime is not null and :new.FinTime is not null  and nvl(:old.InstCount,0) != nvl(:new.InstCount,0) then
      update SO001 set InstCount = nvl(InstCount,0) - nvl(:old.InstCount,0) + nvl(:new.InstCount,0)
        where CustId = :new.CustId;
 
    -- 更新"已完工資料被反為未完工且裝機台數!=0"才異動InstCount
    elsif :old.FinTime is not null and :new.FinTime is null and nvl(:old.InstCount,0)!=0 then
      update SO001 set InstCount = nvl(InstCount,0) - nvl(:old.InstCount,0)
        where CustId = :old.CustId;
    end if;
  end;
/
 
-- 刪除
create or replace trigger TR_SO001_InstCount3
  after delete on SO007
  for each row
  begin
    if nvl(:old.InstCount,0)!=0 then
      update SO001 set InstCount = nvl(InstCount,0) - nvl(:old.InstCount,0)
        where CustId = :old.CustId;
    end if;  
  end;
/
