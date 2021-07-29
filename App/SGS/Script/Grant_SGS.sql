/*
  檔名: Grant_SGS.sql
  @D:\App\SGS\Script\Grant_SGS.sql
  Date: 2004.03.04
*/

-- 1. grant 權限 => 要用 EMCCT(因為 EMCCT 是 owner) 來執行以下之指令
GRANT DELETE ON EMCCT.SO451 TO SGSCOM;
GRANT INSERT ON EMCCT.SO451 TO SGSCOM;
GRANT SELECT ON EMCCT.SO451 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO451 TO SGSCOM;

GRANT DELETE ON EMCCT.SO452 TO SGSCOM;
GRANT INSERT ON EMCCT.SO452 TO SGSCOM;
GRANT SELECT ON EMCCT.SO452 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO452 TO SGSCOM;

GRANT DELETE ON EMCCT.SO453 TO SGSCOM;
GRANT INSERT ON EMCCT.SO453 TO SGSCOM;
GRANT SELECT ON EMCCT.SO453 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO453 TO SGSCOM;

GRANT DELETE ON EMCCT.SO454 TO SGSCOM;
GRANT INSERT ON EMCCT.SO454 TO SGSCOM;
GRANT SELECT ON EMCCT.SO454 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO454 TO SGSCOM;

GRANT DELETE ON EMCCT.SO455 TO SGSCOM;
GRANT INSERT ON EMCCT.SO455 TO SGSCOM;
GRANT SELECT ON EMCCT.SO455 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO455 TO SGSCOM;

GRANT DELETE ON EMCCT.SO456 TO SGSCOM;
GRANT INSERT ON EMCCT.SO456 TO SGSCOM;
GRANT SELECT ON EMCCT.SO456 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO456 TO SGSCOM;

GRANT DELETE ON EMCCT.SO457 TO SGSCOM;
GRANT INSERT ON EMCCT.SO457 TO SGSCOM;
GRANT SELECT ON EMCCT.SO457 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO457 TO SGSCOM;

GRANT DELETE ON EMCCT.SO458 TO SGSCOM;
GRANT INSERT ON EMCCT.SO458 TO SGSCOM;
GRANT SELECT ON EMCCT.SO458 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO458 TO SGSCOM;

GRANT DELETE ON EMCCT.SO459 TO SGSCOM;
GRANT INSERT ON EMCCT.SO459 TO SGSCOM;
GRANT SELECT ON EMCCT.SO459 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO459 TO SGSCOM;

GRANT DELETE ON EMCCT.SO460 TO SGSCOM;
GRANT INSERT ON EMCCT.SO460 TO SGSCOM;
GRANT SELECT ON EMCCT.SO460 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO460 TO SGSCOM;

GRANT DELETE ON EMCCT.SO461 TO SGSCOM;
GRANT INSERT ON EMCCT.SO461 TO SGSCOM;
GRANT SELECT ON EMCCT.SO461 TO SGSCOM;
GRANT UPDATE ON EMCCT.SO461 TO SGSCOM;


GRANT SELECT ON EMCCT.CD024 TO SGSCOM;
GRANT SELECT ON EMCCT.CD024A TO SGSCOM;
GRANT SELECT ON EMCCT.CD022 TO SGSCOM;
GRANT SELECT ON EMCCT.CD034 TO SGSCOM;
GRANT SELECT ON EMCCT.CD024B TO SGSCOM;
GRANT SELECT ON EMCCT.SO001 TO SGSCOM;
GRANT SELECT ON EMCCT.SO004 TO SGSCOM;
GRANT SELECT ON EMCCT.SO005 TO SGSCOM;
GRANT SELECT ON EMCCT.SO105 TO SGSCOM;
GRANT SELECT ON EMCCT.SOAC0202 TO SGSCOM;
GRANT SELECT ON EMCCT.SO002 TO SGSCOM;


CREATE SYNONYM SGSCOM.SO001 FOR EMCCT.SO001;
CREATE SYNONYM SGSCOM.SO004 FOR EMCCT.SO004;
CREATE SYNONYM SGSCOM.SO005 FOR EMCCT.SO005;
CREATE SYNONYM SGSCOM.SO105 FOR EMCCT.SO105;
CREATE SYNONYM SGSCOM.SOAC0202 FOR EMCCT.SOAC0202;
CREATE SYNONYM SGSCOM.CD024 FOR EMCCT.CD024;
CREATE SYNONYM SGSCOM.CD024A FOR EMCCT.CD024A;
CREATE SYNONYM SGSCOM.CD024B FOR EMCCT.CD024B;
CREATE SYNONYM SGSCOM.CD022 FOR EMCCT.CD022;
CREATE SYNONYM SGSCOM.CD034 FOR EMCCT.CD034;
CREATE SYNONYM SGSCOM.SO002 FOR EMCCT.SO002;


CREATE SYNONYM SGSCOM.SO451 FOR EMCCT.SO451;
CREATE SYNONYM SGSCOM.SO452 FOR EMCCT.SO452;
CREATE SYNONYM SGSCOM.SO453 FOR EMCCT.SO453;
CREATE SYNONYM SGSCOM.SO454 FOR EMCCT.SO454;
CREATE SYNONYM SGSCOM.SO455 FOR EMCCT.SO455;
CREATE SYNONYM SGSCOM.SO456 FOR EMCCT.SO456;
CREATE SYNONYM SGSCOM.SO457 FOR EMCCT.SO457;
CREATE SYNONYM SGSCOM.SO458 FOR EMCCT.SO458;
CREATE SYNONYM SGSCOM.SO459 FOR EMCCT.SO459;
CREATE SYNONYM SGSCOM.SO460 FOR EMCCT.SO460;
CREATE SYNONYM SGSCOM.SO461 FOR EMCCT.SO461;





GRANT EXECUTE ON EMCCT.SF_ICCARDQUERY TO SGSCOM;
CREATE SYNONYM SGSCOM.SF_ICCARDQUERY FOR EMCCT.SF_ICCARDQUERY;
GRANT EXECUTE ON EMCCT.SF_CAPRODUCTQUERY TO SGSCOM;
CREATE SYNONYM SGSCOM.SF_CAPRODUCTQUERY FOR EMCCT.SF_CAPRODUCTQUERY;
GRANT EXECUTE ON EMCCT.SF_PRODPURCHASEQUERY TO SGSCOM;
CREATE SYNONYM SGSCOM.SF_PRODPURCHASEQUERY FOR EMCCT.SF_PRODPURCHASEQUERY;
GRANT EXECUTE ON EMCCT.SF_ENTITLEMENTQUERY TO SGSCOM;
CREATE SYNONYM SGSCOM.SF_ENTITLEMENTQUERY FOR EMCCT.SF_ENTITLEMENTQUERY;

GRANT EXECUTE ON EMCCT.SP_INSERTSO453 TO SGSCOM;
CREATE SYNONYM SGSCOM.SP_INSERTSO453 FOR EMCCT.SP_INSERTSO453;





GRANT SELECT ON EMCCT.S_SGSMSG_SEQNO TO SGSCOM;
CREATE SYNONYM SGSCOM.S_SGSMSG_SEQNO FOR EMCCT.S_SGSMSG_SEQNO;


GRANT SELECT ON EMCCT.S_SGSORDER_SeqNo TO SGSCOM;
CREATE SYNONYM SGSCOM.S_SGSORDER_SeqNo FOR EMCCT.S_SGSORDER_SeqNo;
















