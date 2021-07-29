 
-- �s�W
create or replace trigger TR_SO001_InstCount1
  after insert on SO007
  for each row
  declare
  v_RefNo Number;
  begin
    select RefNo into v_RefNo from cd005 where CodeNo=:new.InstCode;
    v_RefNo := nvl(v_RefNo,0);
 
    -- �s�W"���u��ƥB�˾��x��!=0"�~����InstCount
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
 
-- �ק�
create or replace trigger TR_SO001_InstCount2
  after update on SO007
  for each row
  declare
  v_RefNo Number;
  begin
    select RefNo into v_RefNo from cd005 where CodeNo=:new.InstCode;
    v_RefNo := nvl(v_RefNo,0);
 
    -- ��s"���u��ƥB�˾��x��!=0"�~����InstCount
    if :old.FinTime is null and :new.FinTime is not null and nvl(:new.InstCount,0)!=0 then
        if v_RefNo=1 or v_RefNo=5 then
            update SO001 set InstCount = nvl(:new.InstCount,0)
                where CustId = :new.CustId;
        else
            update SO001 set InstCount = nvl(InstCount,0) + nvl(:new.InstCount,0)
                where CustId = :new.CustId;
        end if;
    -- ��s"���u��ƥB�˾��x�Ƥ��P"�~����InstCount
    elsif :old.FinTime is not null and :new.FinTime is not null  and nvl(:old.InstCount,0) != nvl(:new.InstCount,0) then
      update SO001 set InstCount = nvl(InstCount,0) - nvl(:old.InstCount,0) + nvl(:new.InstCount,0)
        where CustId = :new.CustId;
 
    -- ��s"�w���u��ƳQ�Ϭ������u�B�˾��x��!=0"�~����InstCount
    elsif :old.FinTime is not null and :new.FinTime is null and nvl(:old.InstCount,0)!=0 then
      update SO001 set InstCount = nvl(InstCount,0) - nvl(:old.InstCount,0)
        where CustId = :old.CustId;
    end if;
  end;
/
 
-- �R��
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
