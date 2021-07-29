
-- �s�W������Ʈ�: Balance ��h�s���B
create or replace trigger TR_SO001_Balance1
  after insert on SO033
  for each row
  begin
    -- �s�W"�������"�~����Balance (���ΦA + :new.RealAmt)
    if :new.RealDate is null or :new.UCCode is not null then
      update SO001 set Balance = Balance - :new.ShouldAmt
        where CustId = :new.CustId;    
    end if;
  end;
/

-- �ק�������Ʈ�: Balance ��h���B�t�B
create or replace trigger TR_SO001_Balance2
  after update on SO033
  for each row
  begin

    -- �����אּ�w��
    if (:old.RealDate is null or :old.UCCode is not null) and :new.RealDate is not null
      and :new.CancelFlag=0 and :new.UCCode is null then
      update SO001 set Balance = Balance + :old.ShouldAmt
	where CustId = :old.CustId; 

    -- �w���אּ����
    elsif (:new.RealDate is null or :new.UCCode is not null) and :old.RealDate is not null
      and :old.CancelFlag=0 and :old.UCCode is null then
      update SO001 set Balance = Balance - :new.ShouldAmt
	where CustId = :old.CustId; 

    -- �����אּ�@�o
    elsif (:old.RealDate is null or :old.UCCode is not null) and :new.RealDate is not null
      and :new.CancelFlag=1 then
      update SO001 set Balance = Balance + :old.ShouldAmt
	where CustId = :old.CustId; 

    -- �����אּ����,���X����B������
    elsif (:old.RealDate is null or :old.UCCode is not null) and 
      (:new.RealDate is null or :new.UCCode is not null) and 
      :old.ShouldAmt != :new.ShouldAmt then
      update SO001 set Balance = Balance + :old.ShouldAmt - :new.ShouldAmt 
	where CustId = :old.CustId; 

    -- �w���אּ�@�o, �w����w�� ���v�TBalance

    end if;
  end;
/


-- �R����, �Y���������, �h�[�^�X����B
create or replace trigger TR_SO001_Balance3
  after delete on SO033
  for each row
  begin
    if :old.RealDate is null then
      update SO001 set Balance = Balance + :old.ShouldAmt where CustId = :old.CustId;
    end if;  
  end;
/
