-- §ó·s
create or replace trigger TR_UpdCD001
  after update on CD001
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc2('CD001', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD002
  after update on CD002
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc2('CD002', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD003
  after update on CD003
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc2('CD003', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD004
  after update on CD004
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD004', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD005
  after update on CD005
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD005', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD006
  after update on CD006
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD006', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD007
  after update on CD007
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD007', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD008
  after update on CD008
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD008', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD009
  after update on CD009
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD009', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD010
  after update on CD010
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD010', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD011
  after update on CD011
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD011', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD012
  after update on CD012
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD012', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD013
  after update on CD013
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD013', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD014
  after update on CD014
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD014', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD015
  after update on CD015
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD015', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD016
  after update on CD016
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD016', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD017
  after update on CD017
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD017', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD018
  after update on CD018
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD018', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD019
  after update on CD019
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD019', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD020
  after update on CD020
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD020', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD021
  after update on CD021
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD021', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD022
  after update on CD022
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD022', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD024
  after update on CD024
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc2('CD024', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD025
  after update on CD025
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD025', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD026
  after update on CD026
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD026', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD031
  after update on CD031
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD031', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD032
  after update on CD032
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD032', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD034
  after update on CD034
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD034', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD039
  after update on CD039
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD039', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD040
  after update on CD040
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc2('CD040', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD041
  after update on CD041
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD041', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCD042
  after update on CD042
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Description != :old.Description then
        v_RetCode := SF_NewDesc1('CD042', :old.CodeNo, :old.Description, :new.Description, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdCM003
  after update on CM003
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.EmpName  != :old.EmpName then
        v_RetCode := SF_NewDesc2('CM003', :old.EmpNo, :old.EmpName, :new.EmpName, v_RetMsg);
     end if;
  end;
/

create or replace trigger TR_UpdSO017
  after update on SO017
  for each row
  declare
  v_RetMsg Varchar2(20);
  v_RetCode Number;
  begin
     v_RetCode:=0;
     if :new.Name != :old.Name then
        v_RetCode := SF_NewDesc2('SO017', :old.MduId, :old.Name, :new.Name, v_RetMsg);
     end if;
  end;
/
