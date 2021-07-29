
---ViewOwner: com/com@catvn

create or replace view v_catvc_so004 as
select SO004.CompCode,
       SO004.CustID,
       SO001.CustName,
       SO004.FaciSNo,
       SO004.SmartCardNo,
       SO004.BuyName,
       SO004.Deposit,
       SO004.Amount,
       SO004.InstDate,
       SO004.SNo,
       SO004.PRDate,
       SO004.PRSNo,
       SO004.FaciName,
       SO004.UpdTime,
       SO001.Tel1,
       so004.oldinstdate,
       decode(so306.useflag, 2, '움ⅱ', null) blacklist,
       so004.dereasonname
  from emcks.so004@catvc,
       emcks.so001@catvc,
       emcks.cd022@catvc,
       emcks.so306@catvc
 where so001.custid = so004.custid
   and so004.facicode = cd022.codeno
   and so004.facisno = so306.hfcmac(+)
   and cd022.refno in (2, 3, 5)
union all
select SO004.CompCode,
       SO004.CustID,
       SO001.CustName,
       SO004.FaciSNo,
       SO004.SmartCardNo,
       SO004.BuyName,
       SO004.Deposit,
       SO004.Amount,
       SO004.InstDate,
       SO004.SNo,
       SO004.PRDate,
       SO004.PRSNo,
       SO004.FaciName,
       SO004.UpdTime,
       SO001.Tel1,
       so004.oldinstdate,
       decode(so306.useflag, 2, '움ⅱ', null) blacklist,
       so004.dereasonname
  from emcpn.so004@catvc,
       emcpn.so001@catvc,
       emcpn.cd022@catvc,
       emcpn.so306@catvc
 where so001.custid = so004.custid
   and so004.facicode = cd022.codeno
   and so004.facisno = so306.hfcmac(+)
   and cd022.refno in (2, 3, 5)
union all
select SO004.CompCode,
       SO004.CustID,
       SO001.CustName,
       SO004.FaciSNo,
       SO004.SmartCardNo,
       SO004.BuyName,
       SO004.Deposit,
       SO004.Amount,
       SO004.InstDate,
       SO004.SNo,
       SO004.PRDate,
       SO004.PRSNo,
       SO004.FaciName,
       SO004.UpdTime,
       SO001.Tel1,
       so004.oldinstdate,
       decode(so306.useflag, 2, '움ⅱ', null) blacklist,
       so004.dereasonname
  from emcnt.so004@catvc,
       emcnt.so001@catvc,
       emcnt.cd022@catvc,
       emcnt.so306@catvc
 where so001.custid = so004.custid
   and so004.facicode = cd022.codeno
   and so004.facisno = so306.hfcmac(+)
   and cd022.refno in (2, 3, 5)
union all
select SO004.CompCode,
       SO004.CustID,
       SO001.CustName,
       SO004.FaciSNo,
       SO004.SmartCardNo,
       SO004.BuyName,
       SO004.Deposit,
       SO004.Amount,
       SO004.InstDate,
       SO004.SNo,
       SO004.PRDate,
       SO004.PRSNo,
       SO004.FaciName,
       SO004.UpdTime,
       SO001.Tel1,
       so004.oldinstdate,
       decode(so306.useflag, 2, '움ⅱ', null) blacklist,
       so004.dereasonname
  from emcncc.so004@catvc,
       emcncc.so001@catvc,
       emcncc.cd022@catvc,
       emcncc.so306@catvc
 where so001.custid = so004.custid
   and so004.facicode = cd022.codeno
   and so004.facisno = so306.hfcmac(+)
   and cd022.refno in (2, 3, 5)
union all
select SO004.CompCode,
       SO004.CustID,
       SO001.CustName,
       SO004.FaciSNo,
       SO004.SmartCardNo,
       SO004.BuyName,
       SO004.Deposit,
       SO004.Amount,
       SO004.InstDate,
       SO004.SNo,
       SO004.PRDate,
       SO004.PRSNo,
       SO004.FaciName,
       SO004.UpdTime,
       SO001.Tel1,
       so004.oldinstdate,
       decode(so306.useflag, 2, '움ⅱ', null) blacklist,
       so004.dereasonname
  from emcfm.so004@catvc,
       emcfm.so001@catvc,
       emcfm.cd022@catvc,
       emcfm.so306@catvc
 where so001.custid = so004.custid
   and so004.facicode = cd022.codeno
   and so004.facisno = so306.hfcmac(+)
   and cd022.refno in (2, 3, 5)




-- VIEW CATVN

create or replace view v_catvn_so004 as
(select so004.compcode, so004.custid, so001.custname, so004.facisno,
           so004.smartcardno, so004.buyname, so004.deposit, so004.amount,
           so004.instdate, so004.sno, so004.prdate, so004.prsno,
           so004.faciname, so004.updtime, so001.tel1, so004.oldinstdate,
           decode( so306.useflag, 2, '움ⅱ', null ) blacklist,
           so004.dereasonname
      from emcct.so004,
           emcct.so001,
           emcct.cd022,
           emcct.so306
     where so001.custid = so004.custid
       and so004.facicode = cd022.codeno
       and so004.facisno = so306.hfcmac(+)
       and cd022.refno in (2, 3, 5))
   union all
   (select so004.compcode, so004.custid, so001.custname, so004.facisno,
           so004.smartcardno, so004.buyname, so004.deposit, so004.amount,
           so004.instdate, so004.sno, so004.prdate, so004.prsno,
           so004.faciname, so004.updtime, so001.tel1, so004.oldinstdate,
           decode( so306.useflag, 2, '움ⅱ', null ) blacklist,
           so004.dereasonname
      from emcuc.so004,
           emcuc.so001,
           emcuc.cd022,
           emcuc.so306
     where so001.custid = so004.custid
       and so004.facicode = cd022.codeno
       and so004.facisno = so306.hfcmac(+)
       and cd022.refno in (2, 3, 5))
   union all
   (select so004.compcode, so004.custid, so001.custname, so004.facisno,
           so004.smartcardno, so004.buyname, so004.deposit, so004.amount,
           so004.instdate, so004.sno, so004.prdate, so004.prsno,
           so004.faciname, so004.updtime, so001.tel1, so004.oldinstdate,
           decode( so306.useflag, 2, '움ⅱ', null ) blacklist,
           so004.dereasonname
      from emcyms.so004,
           emcyms.so001,
           emcyms.cd022,
           emcyms.so306
     where so001.custid = so004.custid
       and so004.facicode = cd022.codeno
       and so004.facisno = so306.hfcmac(+)
       and cd022.refno in (2, 3, 5))
   union all
   (select so004.compcode, so004.custid, so001.custname, so004.facisno,
           so004.smartcardno, so004.buyname, so004.deposit, so004.amount,
           so004.instdate, so004.sno, so004.prdate, so004.prsno,
           so004.faciname, so004.updtime, so001.tel1,so004.oldinstdate,
           decode( so306.useflag, 2, '움ⅱ', null ) blacklist,
           so004.dereasonname
      from emcntp.so004,
           emcntp.so001,
           emcntp.cd022,
           emcntp.so306
     where so001.custid = so004.custid
       and so004.facicode = cd022.codeno
       and so004.facisno = so306.hfcmac(+)
       and cd022.refno in (2, 3, 5))
   union all
   (select so004.compcode, so004.custid, so001.custname, so004.facisno,
           so004.smartcardno, so004.buyname, so004.deposit, so004.amount,
           so004.instdate, so004.sno, so004.prdate, so004.prsno,
           so004.faciname, so004.updtime, so001.tel1,so004.oldinstdate,
           decode( so306.useflag, 2, '움ⅱ', null ) blacklist,
           so004.dereasonname
      from emckp.so004,
           emckp.so001,
           emckp.cd022,
           emckp.so306
     where so001.custid = so004.custid
       and so004.facicode = cd022.codeno
       and so004.facisno = so306.hfcmac(+)
       and cd022.refno in (2, 3, 5))
   union all
   (select so004.compcode, so004.custid, so001.custname, so004.facisno,
           so004.smartcardno, so004.buyname, so004.deposit, so004.amount,
           so004.instdate, so004.sno, so004.prdate, so004.prsno,
           so004.faciname, so004.updtime, so001.tel1,so004.oldinstdate,
           decode( so306.useflag, 2, '움ⅱ', null ) blacklist,
           so004.dereasonname
      from emcdw.so004,
           emcdw.so001,
           emcdw.cd022,
           emcdw.so306
     where so001.custid = so004.custid
       and so004.facicode = cd022.codeno
       and so004.facisno = so306.hfcmac(+)
       and cd022.refno in (2, 3, 5))
   union all
   (select so004.compcode, so004.custid, so001.custname, so004.facisno,
           so004.smartcardno, so004.buyname, so004.deposit, so004.amount,
           so004.instdate, so004.sno, so004.prdate, so004.prsno,
           so004.faciname, so004.updtime, so001.tel1, so004.oldinstdate,
           decode( so306.useflag, 2, '움ⅱ', null ) blacklist,
           so004.dereasonname
      from emcnty.so004,
           emcnty.so001,
           emcnty.cd022,
           emcnty.so306
     where so001.custid = so004.custid
       and so004.facicode = cd022.codeno
       and so004.facisno = so306.hfcmac(+)
       and cd022.refno in (2, 3, 5))



-- VIEW SDTCMIS
create or replace view v_sdtcmis_so004 as
(select SO004.CompCode,SO004.CustID,SO001.CustName,SO004.FaciSNo,
        SO004.SmartCardNo,SO004.BuyName,SO004.Deposit,SO004.Amount,
        SO004.InstDate,SO004.SNo,SO004.PRDate,SO004.PRSNo,SO004.FaciName,
        SO004.UpdTime,SO001.Tel1, SO004.OLDINSTDATE, null as blacklist, SO004.DEREASONNAME
   from emctc.so004@sdtcmis,
        emctc.so001@sdtcmis,
        emctc.cd022@sdtcmis
  where so001.custid = so004.custid
    and so004.facicode = cd022.codeno
    and cd022.refno in (2,3,5)
 union all
select SO004.CompCode,SO004.CustID,SO001.CustName,SO004.FaciSNo,
        SO004.SmartCardNo,SO004.BuyName,SO004.Deposit,SO004.Amount,
        SO004.InstDate,SO004.SNo,SO004.PRDate,SO004.PRSNo,SO004.FaciName,
        SO004.UpdTime,SO001.Tel1, SO004.OLDINSTDATE,
        decode( so306.useflag, 2, '움ⅱ', null ) blacklist,
        SO004.DEREASONNAME
   from emctc_cm.so004,
        emctc_cm.so001,
        emctc_cm.cd022,
        emctc_cm.so306
  where so001.custid = so004.custid
    and so004.facicode = cd022.codeno
    and so004.facisno = so306.hfcmac(+)
    and cd022.refno in (2,3,5))