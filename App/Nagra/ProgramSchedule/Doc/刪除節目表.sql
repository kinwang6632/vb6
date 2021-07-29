PL/SQL Developer Test script 3.0
47
-- Created on 2006/3/29 by NICK 
declare 
  -- Local variables here
  a155Count integer;
  a155ACount integer;
  aDate varchar2(14);
begin

  aDate := '20060301';

  -- Test statements here
  dbms_output.put_line( '刪除節目表' );
  dbms_output.put_line( '1. SO159 ( Wrapper - PPV )' );
  dbms_output.put_line( '2. SO155/SO155A ( PPV )' );
  dbms_output.put_line( '3. SO158 ( Dex  )' );
  dbms_output.put_line( '4. SO161 ( Wrapper - SUB )' );
  dbms_output.put_line( '5. SO160 ( 暫存處理 )' );
  dbms_output.put_line( '' );
  dbms_output.put_line( '---------------------------------' );
  dbms_output.put_line( '日期:' || aDate );
  
  delete from so159 where eventbegintime >= aDate;
  dbms_output.put_line( 'SO159共刪除了' || to_char( sql%rowcount ) || '筆' );  

  for c1 in ( select unique productid from so158 where salebegintime >= aDate ) loop

      delete from so155 where productid = c1.productid;
      a155Count := ( nvl( a155Count, 0 ) + sql%rowcount );
      
      delete from so155a where productid = c1.productid;
      a155ACount := ( nvl( a155ACount, 0 ) + sql%rowcount );  
  
  end loop;
  
  dbms_output.put_line( 'SO155共刪除了' || to_char( a155Count ) || '筆' );
  dbms_output.put_line( 'SO155A共刪除了' || to_char( a155ACount ) || '筆' );  
  
  delete from so158 where salebegintime >= aDate;
  dbms_output.put_line( 'SO158共刪除了' || to_char( sql%rowcount ) || '筆' );  

  delete from so161 where eventbegintime >= aDate;
  dbms_output.put_line( 'SO161共刪除了' || to_char( sql%rowcount ) || '筆' );

  delete from so160;
  dbms_output.put_line( 'SO160共刪除了' || to_char( sql%rowcount ) || '筆' );

end;
0
0
