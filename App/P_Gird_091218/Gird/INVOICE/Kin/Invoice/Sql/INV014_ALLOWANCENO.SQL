prompt Created on 2006/10/16 by Nick
prompt --------------------------------------------------------------------------------
prompt 以下為 Script 變更之處
prompt 1.建立 sequence  S_INV014_ALLOWANCE
prompt 2.INV014 新增欄位 ALLOWANCENO VARCHAR2(12)
prompt 3.INV014.ALLOWANCENO 填補序號, 編碼規則: 年+月+6碼序號
prompt 4.變更 INV014 的 primary key 為 identifyid1, identifyid2, compid, allowanceno
prompt 5.INV011 新增 'B04.自動產生折讓作業'
prompt --------------------------------------------------------------------------------

drop sequence S_INV014_ALLOWANCE;
create sequence S_INV014_ALLOWANCE minvalue 1 maxvalue 999999 start with 1 increment by 1 cache 20 cycle;

alter table inv014 drop constraint pk_inv014 cascade;

alter table inv014 add allowanceno varchar2(12);

declare
  -- Local variables here
  aSeq integer;
begin
  --- 先分出日期群組  ( 年月 )
  for c1 in ( select to_char( upttime, 'YYYYMM' ) gdstime from inv014 group by to_char( upttime, 'YYYYMM' ) )
  loop

      --- 取出該月份的折讓資料
      for c2 in ( select a.*, a.rowid from inv014 a where to_char( a.upttime, 'YYYYMM' ) = c1.gdstime
                   order by a.upttime )
      loop

          select S_INV014_ALLOWANCE.Nextval into aSeq from dual;

          update inv014 set allowanceno = ( to_char( upttime, 'YYYYMM' ) || lpad( aSeq, 6, '0' ) )
            where rowid = c2.rowid;

      end loop;

  end loop;
end;
/

alter table inv014 add constraint pk_inv014 primary key (identifyid1, identifyid2, compid, allowanceno);

insert into inv011 values ( 1, 0, 'B04', '自動產生折讓作業', sysdate, user );
insert into inv011 values ( 1, 0, 'B05', '零稅率銷售額資料匯檔作業', sysdate, user );
