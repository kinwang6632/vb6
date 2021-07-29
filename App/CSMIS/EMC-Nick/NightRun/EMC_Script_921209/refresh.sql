 -- ks 

connect emcks/emcks@catv0
spool c:\temp\snapks.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('emcks.emcks1');
execute DBMS_REFRESH.REFRESH('emcks.emcks2');
execute DBMS_REFRESH.REFRESH('emcks.emcks3');
execute DBMS_REFRESH.REFRESH('emcks.emcks4');
execute DBMS_REFRESH.REFRESH('emcks.emcks5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--pn
connect emcpn/emcpn@catv0
spool c:\temp\snappn.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('emcpn.emcpn1');
execute DBMS_REFRESH.REFRESH('emcpn.emcpn2');
execute DBMS_REFRESH.REFRESH('emcpn.emcpn3');
execute DBMS_REFRESH.REFRESH('emcpn.emcpn4');
execute DBMS_REFRESH.REFRESH('emcpn.emcpn5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--nt
connect emcnt/emcnt@catv0
spool c:\temp\snapnt.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('emcnt.emcnt1');
execute DBMS_REFRESH.REFRESH('emcnt.emcnt2');
execute DBMS_REFRESH.REFRESH('emcnt.emcnt3');
execute DBMS_REFRESH.REFRESH('emcnt.emcnt4');
execute DBMS_REFRESH.REFRESH('emcnt.emcnt5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--ncc
connect emcncc/emcncc@catv0
spool c:\temp\snapncc.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('emcncc.emcncc1');
execute DBMS_REFRESH.REFRESH('emcncc.emcncc2');
execute DBMS_REFRESH.REFRESH('emcncc.emcncc3');
execute DBMS_REFRESH.REFRESH('emcncc.emcncc4');
execute DBMS_REFRESH.REFRESH('emcncc.emcncc5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--fm
connect emcfm/emcfm@catv0
spool c:\temp\snapfm.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('emcfm.emcfm1');
execute DBMS_REFRESH.REFRESH('emcfm.emcfm2');
execute DBMS_REFRESH.REFRESH('emcfm.emcfm3');
execute DBMS_REFRESH.REFRESH('emcfm.emcfm4');
execute DBMS_REFRESH.REFRESH('emcfm.emcfm5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--ct
connect emcct/emcct@catv0
spool c:\temp\snapct.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('emcct.emcct1');
execute DBMS_REFRESH.REFRESH('emcct.emcct2');
execute DBMS_REFRESH.REFRESH('emcct.emcct3');
execute DBMS_REFRESH.REFRESH('emcct.emcct4');
execute DBMS_REFRESH.REFRESH('emcct.emcct5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--uc
connect emcuc/emcuc@catv0
spool c:\temp\snapuc.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('emcuc.emcuc1');
execute DBMS_REFRESH.REFRESH('emcuc.emcuc2');
execute DBMS_REFRESH.REFRESH('emcuc.emcuc3');
execute DBMS_REFRESH.REFRESH('emcuc.emcuc4');
execute DBMS_REFRESH.REFRESH('emcuc.emcuc5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--tc
connect emctc/emctc@catv0
spool c:\temp\snaptc.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('emctc.emctc1');
execute DBMS_REFRESH.REFRESH('emctc.emctc2');
execute DBMS_REFRESH.REFRESH('emctc.emctc3');
execute DBMS_REFRESH.REFRESH('emctc.emctc4');
execute DBMS_REFRESH.REFRESH('emctc.emctc5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--st
connect emcst/emcst@catv0
spool c:\temp\snapst.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('emcst.emcst1');
execute DBMS_REFRESH.REFRESH('emcst.emcst2');
execute DBMS_REFRESH.REFRESH('emcst.emcst3');
execute DBMS_REFRESH.REFRESH('emcst.emcst4');
execute DBMS_REFRESH.REFRESH('emcst.emcst5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--msoyms
connect msoyms/msoyms@catv0
spool c:\temp\snapyms.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('msoyms.msoyms1');
execute DBMS_REFRESH.REFRESH('msoyms.msoyms2');
execute DBMS_REFRESH.REFRESH('msoyms.msoyms3');
execute DBMS_REFRESH.REFRESH('msoyms.msoyms4');
execute DBMS_REFRESH.REFRESH('msoyms.msoyms5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--msontp
connect msontp/msontp@catv0
spool c:\temp\snapntp.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('msontp.msontp1');
execute DBMS_REFRESH.REFRESH('msontp.msontp2');
execute DBMS_REFRESH.REFRESH('msontp.msontp3');
execute DBMS_REFRESH.REFRESH('msontp.msontp4');
execute DBMS_REFRESH.REFRESH('msontp.msontp5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--msokp
connect msokp/msokp@catv0
spool c:\temp\snapkp.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('msokp.msokp1');
execute DBMS_REFRESH.REFRESH('msokp.msokp2');
execute DBMS_REFRESH.REFRESH('msokp.msokp3');
execute DBMS_REFRESH.REFRESH('msokp.msokp4');
execute DBMS_REFRESH.REFRESH('msokp.msokp5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off
--msolk
connect msolk/msolk@catv0
spool c:\temp\snaplk.log
prompt 起始時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
execute DBMS_REFRESH.REFRESH('msolk.msolk1');
execute DBMS_REFRESH.REFRESH('msolk.msolk2');
execute DBMS_REFRESH.REFRESH('msolk.msolk3');
execute DBMS_REFRESH.REFRESH('msolk.msolk4');
execute DBMS_REFRESH.REFRESH('msolk.msolk5');
prompt 完成時間
select to_char(sysdate,'YYYY/MM/DD HH24:MI:SS') TIME from dual;
spool off