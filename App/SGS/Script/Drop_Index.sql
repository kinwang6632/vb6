/*
  Drop indexes before test data been loaded

  By: Pierre
  Date: 2004.04.07
	2004.04.13 Add 1 new index for SO004
	2004.04.16 Add 1 new index for SO004, add index for SO005 & SOAC0202
*/

PROMPT INDEX NAME : I_SO004_OrderNo
Drop INDEX I_SO004_OrderNo;

PROMPT INDEX NAME : I_SO004_SSBFP
Drop INDEX I_SO004_SSBFP;

PROMPT INDEX NAME : I_SO004_CustId_FaciCode
Drop INDEX I_SO004_CustId_FaciCode;

PROMPT INDEX NAME : I_SO004_CvtId
Drop INDEX I_SO004_CvtId;

PROMPT INDEX NAME : I_SO004_FaciSNo
Drop INDEX I_SO004_FaciSNo;

PROMPT INDEX NAME : I_SO004_CSS
Drop INDEX I_SO004_CSS;

PROMPT INDEX NAME : I_SO004_MediaCode
Drop INDEX I_SO004_MediaCode;

PROMPT INDEX NAME : I_SO004_PRDate
Drop INDEX I_SO004_PRDate;

PROMPT INDEX NAME : I_SO004_FaciCode
Drop INDEX I_SO004_FaciCode;

PROMPT INDEX NAME : I_SO004_BuyCode
Drop INDEX I_SO004_BuyCode;

PROMPT INDEX NAME : I_SO004_InstDate
Drop INDEX I_SO004_InstDate;

PROMPT INDEX NAME : I_SO004_PRSNO
Drop INDEX I_SO004_PRSNO;

PROMPT INDEX NAME : I_SO004_CheckNo
Drop INDEX I_SO004_CheckNo;

PROMPT INDEX NAME : I_SO004_MACAddress
Drop INDEX I_SO004_MACAddress;

PROMPT INDEX NAME : I_SO004_SmartCardNo
Drop INDEX I_SO004_SmartCardNo;

PROMPT INDEX NAME : I_SO004_SGS_1
Drop INDEX I_SO004_SGS_1;


PROMPT INDEX NAME : I_SO005_CustId
Drop INDEX I_SO005_CustId;

PROMPT INDEX NAME : I_SO005_CustIdCvtIdChCode
Drop INDEX I_SO005_CustIdCvtIdChCode;

PROMPT INDEX NAME : I_SO005_CSC
Drop INDEX I_SO005_CSC;


PROMPT INDEX NAME : I_SOAC0202_ICCSNO
Drop INDEX I_SOAC0202_ICCSNO;

PROMPT INDEX NAME : I_SOAC0202_STBSNO
Drop INDEX I_SOAC0202_STBSNO;


PROMPT INDEX NAME : I_SO451_HandleTime
Drop INDEX I_SO451_HandleTime;

PROMPT INDEX NAME : I_SO453_HandleTime
Drop INDEX I_SO453_HandleTime;

PROMPT INDEX NAME : I_SO455_HandleTime
Drop INDEX I_SO455_HandleTime;

PROMPT INDEX NAME : I_SO457_HandleTime
Drop INDEX I_SO457_HandleTime;

PROMPT Ok !!!


