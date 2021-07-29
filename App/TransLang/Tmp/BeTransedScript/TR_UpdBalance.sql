
-- 新增應收資料時: Balance 減去新金額
create or replace trigger TR_SO001_Balance1
  after insert on SO033
  for each row
  begin
    -- 新增"未收資料"才異動Balance (不用再 + :new.RealAmt)
    if :new.RealDate is null or :new.UCCode is not null then
      update SO001 set Balance = Balance - :new.ShouldAmt
        where CustId = :new.CustId;    
    end if;
  end;
/

-- 修改應收資料時: Balance 減去金額差額
create or replace trigger TR_SO001_Balance2
  after update on SO033
  for each row
  begin

    -- 未收改為已收
    if (:old.RealDate is null or :old.UCCode is not null) and :new.RealDate is not null
      and :new.CancelFlag=0 and :new.UCCode is null then
      update SO001 set Balance = Balance + :old.ShouldAmt
	where CustId = :old.CustId; 

    -- 已收改為未收
    elsif (:new.RealDate is null or :new.UCCode is not null) and :old.RealDate is not null
      and :old.CancelFlag=0 and :old.UCCode is null then
      update SO001 set Balance = Balance - :new.ShouldAmt
	where CustId = :old.CustId; 

    -- 未收改為作廢
    elsif (:old.RealDate is null or :old.UCCode is not null) and :new.RealDate is not null
      and :new.CancelFlag=1 then
      update SO001 set Balance = Balance + :old.ShouldAmt
	where CustId = :old.CustId; 

    -- 未收改為未收,但出單金額有改變
    elsif (:old.RealDate is null or :old.UCCode is not null) and 
      (:new.RealDate is null or :new.UCCode is not null) and 
      :old.ShouldAmt != :new.ShouldAmt then
      update SO001 set Balance = Balance + :old.ShouldAmt - :new.ShouldAmt 
	where CustId = :old.CustId; 

    -- 已收改為作廢, 已收改已收 不影響Balance

    end if;
  end;
/


-- 刪除時, 若為未收資料, 則加回出單金額
create or replace trigger TR_SO001_Balance3
  after delete on SO033
  for each row
  begin
    if :old.RealDate is null then
      update SO001 set Balance = Balance + :old.ShouldAmt where CustId = :old.CustId;
    end if;  
  end;
/
