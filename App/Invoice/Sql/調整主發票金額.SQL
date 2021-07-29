-- Created on 2006/11/10 by NICK
-- 叫把σ粂猭,秸俱祇布郎ず, 笆传秨祇布祇布肂ぃタ絋戈
declare
  -- Local variables here
  aMainSaleAmount inv007.mainsaleamount%type;
  aMainTaxAmount inv007.maintaxamount%type;
  aMainInvAmount inv007.maininvamount%type;
  aPreMainInvId  inv007.invid%type;
  aApplyChange boolean;
begin


  for c1 in (  select invdate, count(1) from inv007 where compid = "" group by invdate )
  loop

      aPreMainInvId := null;
      aMainSaleAmount := 0;
      aMainTaxAmount := 0;
      aMainInvAmount := 0;

      aApplyChange := false;

      for c2 in (  select a.* from inv007 a,
                     ( select unique maininvid from inv007
                        where compid = ""
                          and invdate = c1.invdate
                          and maininvid <> invid
                          and ( ( invamount = maininvamount ) or ( saleamount = mainsaleamount ) ) ) b
                    where a.maininvid = b.maininvid
                    order by a.maininvid, a.invid )
      loop

         if ( aPreMainInvId is null ) then aPreMainInvId := c2.maininvid; end if;

         if ( c2.maininvid <> aPreMainInvId ) then


            dbms_output.put_line( '祇布腹絏:' || aPreMainInvId || ', ' ||
                                  '祇布肂=' || to_char( aMainInvAmount ) || ', ' ||
                                  '祇布綪扳肂=' || to_char( aMainSaleAmount ) || ', ' ||
                                  '祇布祙肂=' || to_char( aMainTaxAmount ) );

            update inv007
               set mainsaleamount = aMainSaleAmount,
                   maintaxamount = aMainTaxAmount,
                   maininvamount = aMainInvAmount
             where maininvid = aPreMainInvId;

            aMainSaleAmount := 0;
            aMainTaxAmount := 0;
            aMainInvAmount := 0;
            aPreMainInvId := c2.maininvid;

         end if;

         aMainSaleAmount := ( aMainSaleAmount + c2.saleamount );
         aMainTaxAmount := ( aMainTaxAmount + c2.taxamount );
         aMainInvAmount := ( aMainInvAmount + c2.Invamount );

         aApplyChange := true;

      end loop;

      if ( aApplyChange ) then

         dbms_output.put_line( '祇布腹絏:' || aPreMainInvId || ', ' ||
                               '祇布肂=' || to_char( aMainInvAmount ) || ', ' ||
                               '祇布綪扳肂=' || to_char( aMainSaleAmount ) || ', ' ||
                               '祇布祙肂=' || to_char( aMainTaxAmount ) );
         update inv007
            set mainsaleamount = aMainSaleAmount,
                maintaxamount = aMainTaxAmount,
                maininvamount = aMainInvAmount
           where maininvid = aPreMainInvId;

       end if;

  end loop;
end;
/