declare 
  aStartTime date;
  aEndTime number;
  aCounts integer;
  aInv008Count integer;
  aInv008ACount integer;
begin
  dbms_output.put_line( '�}�l��J�ɶ�: ' || to_char( sysdate, 'HH24:MI:SS' ) );
  -----
  for c1 in ( select to_char( invdate, 'YYYY/MM' ) yearmonth, sum( counts ) records
                from ( select invdate, count(1) counts from inv007 group by invdate ) a group by to_char( invdate, 'YYYY/MM' ) )
  loop
      aStartTime := sysdate;
      -----
      insert /*+ APPEND */ into inv008a ( invid, billid, billiditemno, seq, startdate, enddate, 
        itemid, description, quantity, unitprice, taxamount, saleamount, totalamount, servicetype )
          select b.InvID, b.BillID, b.BillIDItemNo, b.Seq, b.StartDate, b.EndDate, 
                 b.ItemID, b.Description, b.Quantity, b.UnitPrice, b.TaxAmount, b.SaleAmount, b.TotalAmount, b.ServiceType
            from inv007 a, inv008 b where a.invid = b.invid
             and a.invdate between to_date( c1.yearmonth, 'YYYY/MM' ) and last_day( to_date( c1.yearmonth, 'YYYY/MM' ) );
      aCounts := sql%rowcount;       
      commit;
      -----
      aEndTime := ( sysdate - aStartTime ) * ( 1 / 24 / 60 / 60 );
      if ( aEndTime <= 0.001 ) then aEndTime := 1; end if;
      dbms_output.put_line( '�o���~��:' || c1.yearmonth || ', INV008A �@��J' || to_char( aCounts ) ||' �����, ��O���: ' || to_char( aEndTime ) );
  end loop;
  -----
  dbms_output.put_line( '������J�ɶ�: ' || to_char( sysdate, 'HH24:MI:SS' ) );
  dbms_output.put_line( '----------------------------------------------------' );
  -----
  select count(1) into aInv008Count from inv008;
  select count(1) into aInv008ACount from inv008a;
  -----
  dbms_output.put_line( '��������:  INV008: ' || to_char( aInv008Count ) );
  dbms_output.put_line( '          INV008A: ' || to_char( aInv008ACount ) );
  -----
end;
