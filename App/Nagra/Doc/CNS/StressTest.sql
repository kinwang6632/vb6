-- Created on 2009/1/20 by NICK
declare
  -- Local variables here
  aIccNo integer;
  aStbNo varchar2(12);
  aZipCode varchar2(4);
  aBatId varchar2(5);
  aSeqNo integer;
  aProductBeginDate varchar2(8);
  aProductEndDate varchar2(8);
  aNotes varchar2(2000);
begin

  delete from SEND_NAGRA where operator = 'Nick-Batch';

  aIccNo := 2145060000;
  aStbNo := '65536';

  aZipCode := '110';
  aBatId := '25148';


  for aIndex in 2145060000..2145061999
  loop
    aIccNo := ( aIndex + 1 );
    select s_sendnagra_seqno.nextval into aSeqNo from dual;

    insert into SEND_NAGRA
      ( high_level_cmd_id, icc_no, stb_no, cmd_status,
        operator, updtime, zip_code, mis_ird_cmd_data,
        seqno, compcode, resenttimes, stb_flag )
    values
      ( 'A1', aIccNo, aStbNo, 'W',
        'Nick-Batch', sysdate, aZipCode, aBatId,
        aSeqNo, 7, 0, 1 );

  end loop;

  aProductBeginDate := to_char( sysdate, 'YYYYMMDD' );
  aProductEndDate := '20201231';

  aNotes :=
    ( 'B~20~' || aProductBeginDate || '~' || aProductEndDate || ',' ||
      'B~21~' || aProductBeginDate || '~' || aProductEndDate );

  for aIndex in 2145060000..2145061999
  loop
    aIccNo := ( aIndex + 1 );
    select s_sendnagra_seqno.nextval into aSeqNo from dual;

    insert into SEND_NAGRA
      ( high_level_cmd_id, icc_no, stb_no, cmd_status,
        operator, updtime, Notes,
        seqno, compcode, resenttimes, stb_flag )
    values
      ( 'B1', aIccNo, aStbNo, 'W',
        'Nick-Batch', sysdate, aNotes,
        aSeqNo, 7, 0, 1 );

  end loop;

end;
/