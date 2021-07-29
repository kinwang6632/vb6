unit cbResStr;

interface

resourcestring

  SVendor = '�}�լ��';
  SAppTitle = 'SMS Command Gateway';
  SVersion = '����: 5.3 ( Build 5.3.0 )';
  
  SUnknownError = '�������~�C';
  SLicenceNotRegister = '���n��|�����U,�Y�ݦX�k�ϥνгs��%s�C';
  SLicenceExpire = '���n��ϥδ����w�L,�Y�ݨϥνгs��%s�C';
  SLicenceImmediateExpire = '���n��w�Q����ϥ�,�гs��%s�C';
  SLicenceInvalid = '���n����U��Ƥ��X�k�C';

  SInputFormTitle = '�n����U';
  SLicenceNote = '�д��ѤU�C�n��Ǹ�,�ÿ�J���U���X,�H���o���n��X�k�ϥ��v�C';
  SLicenceNote2 = '��J�n��Ǹ��i���͵��U���X�C';
  SLicenceKey = '�n��Ǹ�';
  SRegisterKey = '���U�Ǹ�';
  SInstallDate = '�w�ˤ��';
  SExpireDate = '�ϥΨ����';
  SMacList = 'MAC�Ǹ�';


  SConfigSearchFile = '�j�M�]�w�ɤ��C';
  SConfigSearchCompelete = '�@���%d�ӳ]�w�ɡC';

  SConfigFileNotAssign = '�|�����w�]�w�ɡC';
  SConfigFileNotExists = '���w���]�w��%s���s�b, ���ˬd�ɮצW�٩άO���|�O�_���T�C';
  SConfigFileChange = '�����ܳ]�w��-%s�C';

  SConfigOpenSuccess = '�]�w��%s���J�����C';
  SConfigOpenError = '�]�w��%s���J���~�C��]:%s';

  SConfigSoReadSuccess = '�t�ΥxŪ������,�@���J%d�Өt�Υx�C';
  SConfigSoRead = 'Ū���t�Υx�]�w[%s]�C';
  SConfigSoReadError = '�]�w��,Ū���t�Υx���~�C��]:%s';

  SConfigFeedbackDbRead = 'Ū���^�Ǿ���^�g��Ʈw�]�w�C';
  SConfigFeedbackDbSuccess = '�^�Ǿ���^�g��Ʈw�]�wŪ�������C';

  SConfigSoWriteError = '�]�w��[�t�Υx]�g�J���~�C��]:%s';
  SConfigSoWriteSuccess = '�]�w��[�t�Υx]�g�J�����C';

  SConfigCommonReadError = '�]�w��,Ū���@��]�w���~�C��]:%s';
  SConfigCommonWriteError = '�]�w��,[�@��]�w]�g�J���~�C��]:%s�C';
  SConfigCommonReadSuccess = '�]�w��[�@��]�w]Ū�������C';
  SConfigCommonWritSuccess = '�]�w��[�@��]�w]�g�J�����C';

  SConfigCommonProcessRecordCount = 'Ū���]�w[�C���B�z��Ƶ���]=%d';
  SConfigCommonProcessIPPV = 'Ū���]�w[�B�zIPPV���]=%s';
  SConfigCommonProcessBatch = 'Ū���]�w[�B�z�妸���]=%s';
  SConfigCommonDbMultiThread = 'Ū���]�w[�U�t�Υx�ϥοW�߰����]=%s';
  SConfigCommonDbRetryFrequence = 'Ū���]�w[�t�Υx�_�u�ɭ��s���]=%d';
  SConfigCommonDbWriteErrorWhenSocketFail = 'Ū���]�w[�L�k�ǰe���O��,�N��ƼХܿ��~]=%s';
  SConfigCommonDbResendCount = 'Ū���]�w[��ƭ��e���ƭ���]=%d';
  SConfigCommonBatchOperator = 'Ū���]�w[�妸����ѧO�b��]=%s';
  SConfigCommonCASRetryFrequence = 'Ū���]�w[CAS�_�u�ɭ��s���]=%d';
  SConfigCommonBusyTimeStart = 'Ū���]�w[���L�ɬq�_]=%s';
  SConfigCommonBusyTimeEnd = 'Ū���]�w[���L�ɬq��]=%s';
  SConfigCommonBusyTimeDbReadFrequence = 'Ū���]�w[���L�ɬqŪ����Ƭ��]=%s';
  SConfigCommonNormalTimeDbReadFrequence = 'Ū���]�w[�@��ɬqŪ����Ƭ��]=%s';
  SConfigCommonCASSendErrMax = 'Ū���]�w[�ǰe���O��, �o�ͦh�֦����~�۰��_�u]=%d';
  SConfigCommonCASRecvErrMax = 'Ū���]�w[�������O��, �o�ͦh�֦����~�۰��_�u]=%d';
  SConfigCommonCASCommCheck = 'Ū���]�w[�C�h�ְe�@�� Communication Check]=%d';

  SConfigHighCmdReading = 'Ū���]�w[�������O��Ӫ�]=%s(%s)�C';
  SConfigHighCmdReadError = '�]�w��,Ū��[�������O��Ӫ�]���~�C��]:%s';
  SConfigHighCmdReadIsEmpty = '�]�w��,Ū��[�������O��Ӫ�]���~�C��]:���O��Ӫ��e���šC';
  SConfigHighCmdReadSuccess = '[�������O��Ӫ�]Ū�������C';
  SConfigHighCmdWriteError = '�]�w��[�������O��Ӫ�]�g�J���~�C��]:%s�C';
  SConfigHighCmdWriteSuccess = '�]�w��[�������O��Ӫ�]�g�J�����C';

  SConfigLowCmdReading = 'Ū���]�w[�C�����O�y�z��]=%s(%s)�C';
  SConfigLowCmdReadError = '�]�w��,Ū��[�C�����O�y�z��]���~�C��]:%s';
  SConfigLowCmdReadIsEmpty = '�]�w��,Ū��[�C�����O�y�z��]���~�C��]:���O��Ӫ��e���šC';
  SConfigLowCmdReadSuccess = '[�C�����O�y�z��]Ū�������C';
  SConfigLowCmdWriteError = '�]�w��[�C�����O�y�z��]�g�J���~�C��]:%s�C';
  SConfigLowCmdWriteSuccess = '�]�w��[�C�����O�y�z��]�g�J�����C';

  SConfigCmdErrorReading = 'Ū���]�w[���O���~��Ӫ�]�C';
  SConfigCmdErrorReadSuccess = '[���O���~��Ӫ�]Ū�������C';
  SConfigCmdErrorWriteSuccess = '�]�w��[���O���~��Ӫ�]�g�J�����C';
  SConfigCmdErrorWriteError = '�]�w��[���O���~��Ӫ�]�g�J���~�C��]:%s�C';


  SConfigCmdTransUseSimulator = 'Ū���]�w[�ϥμ�����]=%s�C';
  SConfigCmdTransEnableControlSend = 'Ū���]�w[�ҥζǰe���O]=%s�C';
  SConfigCmdTransEnableFeedbackRecv = 'Ū���]�w[�ҥΦ^�Ǿ���]=%s�C';
  SConfigCmdTransSendCommandDelay = 'Ū���]�w[�ǰe�C�����O������]=%f�C';
  SConfigCmdTransReadSuccess = '���O�ǰe�]�wŪ�������C';
  SConfigCmdTransReadError = '�]�w��,Ū�����O�ǰe�]�w���~�C��]:%s�C';
  SCOnfigCmdTransWriteError = '�]�w��[���O�ǰe]�g�J���~�C��]:%s�C';
  SCOnfigCmdTransWriteSuccess = '�]�w��[���O�ǰe]�g�J�����C';

  SConfigCASReadingIP = 'Ū���]�wCAS�D��[IP]=%s�C';
  SConfigCASReadingSPort = 'Ū���]�wCAS�D��[Control Port]=%d�C';
  SConfigCASReadingRPort = 'Ū���]�wCAS�D��[Feedback Port]=%d�C';

  SConfigCASReadingControlSourceId = 'Ū���]�wCAS�D��[Control Source Id]=%s�C';
  SConfigCASReadingControlDestId = 'Ū���]�wCAS�D��[Control Dest Id]=%s�C';
  SConfigCASReadingControlMopPPId = 'Ū���]�wCAS�D��[Control MopPP Id]=%s�C';

  SConfigCASReadingFeedbackSourceId = 'Ū���]�wCAS�D��[Feedback Source Id]=%s�C';
  SConfigCASReadingFeedbackDestId = 'Ū���]�wCAS�D��[Feedback Dest Id]=%s�C';
  SConfigCASReadingFeedbackMopPPId = 'Ū���]�wCAS�D��[Feedback MopPP Id]=%s�C';

  SConfigCASReadControlTransId = 'Ū���]�wCAS�D��[Feedback Trans Id]=%s�C';
  SConfigCASReadFeedbackTransId = 'Ū���]�wCAS�D��[Feedback Trans Id]=%s�C';
  SConfigCASReadMailId = 'Ū���]�wCAS�D��[Mail Id]=%s�C';
  SConfigCASReadCmdMaxCounter = 'Ū���]�wCAS�D��[�C�����O�ǰe�p�ƾ��W��:]=%d�C';

  SConfigCASReadSuccess = 'CAS�D���]�wŪ�������C';
  SConfigCASWriteError = '�]�w�ɼg�J[CAS�D��]���~�C��]:%s';
  SConfigCASWritSuccess = '�]�w��[CAS�D��]�]�w�g�J�����C';

  SConfigNoneSet = '���]�w';

  SConfigBuildSoTreeStart = '�إߨt�Υx....';
  SConfigBuildSoTree = '�إߨt�Υx[%s]�C';
  SConfigBuildSoTreeEnd = '�t�Υx�إߧ����C';

  SAutoExec = '�w�ҥΦ۰ʰ���,�y��{���N�ۦ�ҰʡC';
  SNoneAutoExec = '�۰ʰ��楼�ҥ�,�Ф�ʰ���C';
  SUnknowExec = '���ҥ�%s��%s,�]�w�ɳ]�w���~�C';
  SExecMode = '�w�]�w�ҥ�%s';

  SDbConnectStart = '�U�t�Υx��Ʈw�s�u��....';
  SDbConnectSuccess = '�t�Υx[%s]��Ʈw�s�������C';
  SDbConnectError = '�t�Υx[%s]��Ʈw�s�����ѡC��]:%s�C%d���N�۰ʭ��s�s���C';
  SDbConnectEnd = '�t�Υx��Ʈw�s�u�����C';

  SDbDisConnectStart = '�U�t�Υx��Ʈw���u���C';
  SDbDisConnectSuccess = '�t�Υx[%s]��Ʈw���u�����C';
  SDbDisConnectError = '�t�Υx[%s]��Ʈw���u���ѡC��]:%s';
  

  SFeedbackDbConnectStart = '[�^�Ǿ���]��Ʈw�s�u��....';
  SFeedbackDbConnectSuccess = '[�^�Ǿ���]��Ʈw�s�������C';
  SFeedbackDbConnectError = '[�^�Ǿ���]��Ʈw�s�����ѡC��]:%s�C%d���N�۰ʭ��s�s���C';
  SFeedbackDbbConnectEnd = '[�^�Ǿ���]��Ʈw�s�u�����C';

  SFeedbackDbDisConnectStart = '[�^�Ǿ���]��Ʈw���u���C';
  SFeedbackDbDisConnectSuccess = '[�^�Ǿ���]��Ʈw���u�����C';
  SFeedbackDbDisConnectError = '[�^�Ǿ���]��Ʈw���u���ѡC��]:%s�C';
  SFeedbackDbWriteDataError = '[�^�Ǿ���]�^�Ǹ�Ƽg�J���ѡC��]:%s�C';
  SFeedbackDbWriteDataDone = '[�^�Ǿ���]�^�Ǹ�Ƽg�J����,�@�g�J%d���C';
  SFeedbackDbWriteDataError2 = '[�^�Ǿ���]�^�Ǹ�Ƽg�J����,�o��%d���C';

  SFeedbackDbRertyConnect = '[�^�Ǿ���]��Ʈw���s�s���C';
  SFeedbackDbHasProblem = '[�^�Ǿ���]��Ʈw�g�J�w���Ѧh��,%d���N�۰ʭ��s�إ߸�Ʈw�s���C';

  SDbGetDataError = '�t�Υx[%s]�s���ӷ���ƥ��ѡC��]:%s';
  SDbHasProblem = '�t�Υx[%s]��Ʈw�s���w���Ѧh��,%d���N�۰ʭ��s�إ߸�Ʈw�s���C';
  SDbRertyConnect = '�t�Υx[%s]��Ʈw���s�s���C';
  SDbAddToControlSendList = '�t�Υx[%s]�ݳB�z���O%d��,�w�g�J��C���Զǰe�C';
  SDbScanSo = '�t�Υx[%s],�L�ݳB�z��ơC';
  SDbCheckSendListTooMany = '�ݳB�z���O��C�w�W�X���O��Ǽ�[%d/%d],�t�Υx��ƽ��߱N���ԤU���B�z�C';

  SDbWriteAlreadySendError = '�t�Υx[%s]�^�g���O�w�ǰe����,���O:%s,�Ǹ�:%s�C��]:%s�C';
  SDbWriteAlreadySendCount = '�t�Υx[%s]�^�g���O���A[�B�z��],�����B�z%d���C';
  SDbWriteNotSendCount = '�t�Υx[%s]�^�g���O���A[�L�k�ǰe],�����B�z%d���C';
  SDbWriteAckError = '�t�Υx[%s]�^�g���O�^������,���O:%s,�Ǹ�:%s�C��]:%s�C';
  SDbWriteAckCount = '�t�Υx[%s]�^�g���O���A[�w�B�z],��������%d���C';
  SDbWriteLogCount = '�t�Υx[%s]���O�O����x�g�J����,������x�O��%d���ơC';


  SControlSendManager = '�ǰe���O��C';
  SControlSendDoneManager = '�w�ǰe���O��C';
  SControlRecvManager = '���O�^����C';
  SControlSendLogManager = '���O��x��C';
  SCleanupCommandManagerError = '�M��[%s]�o�Ϳ��~�C��]:%s�C';
  SCleanupCommandManagerSuccess =  '�M��[%s]�����C';


  SThreadStart = '�ҩl%s������C';
  SThreadStartError = '�ҩl%s��������ѡC��]:%s�C';
  SThreadReady = '%s������N��,�ݩR���C';
  SThreadStop = '���b����%s������C';
  SThreadStopError = '����%s��������ѡC��]:%s�C';
  SThreadEnd = '%s���������w�����C';

  SThreadHasProblem = '%s������PCAS�D���s������, %d���N�۰ʭ��s�إ߳s���C';
  SDbThreadRunDelay = '�t�Υx��Ʈw����s��, ����CAS�D���s���C';

  SControlSend = '[�ǰe���O]';
  SControlRecv = '[�����^��]';
  SSoDb = '[�t�Υx��Ʈw]';
  SFeedbackDb = '[�^�Ǿ����Ʈw]';
  SFeedbackSend = '[�^�Ǧ^��]';
  SFeedbackRecv = '[�^�Ǹ��]';
  SGuiCleanup = '[�M���e����ܫ��O]';
  SReleaseControl = '[����D��]';

  SRunStop = '%s�w�������C';

  SSocketBeginConnection = '%s�D�� ( %s : %d ) �s�����C';
  SSocketConnectionFailure = '%s�D���s�����ѡC��]:TCP/IP�s�u���ѡC';
  SSocketConnectionReject = '%s�D���s�����ѡC��]:NAGRA�ڵ��s�u�C';
  SSocketConnectionPostpend = '%s�D���s�����ѡC��]:NAGRA�t���c���C';
  SSocketConnectionError = '%s�D���s�����ѡC��]:%s�C';
  SSocketConnectionSuccess = '%s�D���s�������C';
  SSocketThreadError = '�D��%s��������ѡC��]:%s�C';

  SControlSendEndConnection = '[�ǰe���O]�P�D�� ( %s : %d ) ���u���C';
  SControlSendEndConnectionSuccess = '[�ǰe���O]�D�����u�����C';
  SControlSendEndConnectionError = '[�ǰe���O]�P�D�����u���ѡC��]:%s�C';

  SSendCommandError = '%s����, �������O:%s, �C�����O:%s�C��]:%s�C';
  SSendErrors = '%s�w����%d���H�W,�t�αN����ǰe���O,�òפ�D���s���C';
  SCommunicationCheckError = '�D�� ( %s : %d ) �ڵ�%s,�t�αN����%s,�òפ�s���C';
  SWaitRetry = '%s%d���N�۰ʭ��s�إ߳s���C';
  SEnsureCandSendCommand = '%s�T�{���C';
  SEnsureOK = '%s�T�{����,�i�H�}�l%s�C';
  SEnsureError = '%s�T�{����,%s�Q�ڵ�,�t�αN����ǰe���O,�òפ�s���C';
  SControlSendRecord = '%s�w�B�z���ݶǰe��C���O,�������O%d��,�C�����O%d��';
  SControlSendCasMaxCounterWarning1 = '%s�w�F��C�����O�ǰe�p�ƾ��W��,�N�۰ʭ��C���O�ǰe�t�v,�w�קK CAS �D���t���L�q�C';
  SFeedbackSendRecord = '%s�w�B�z�^����C���O,�@�p%d���C';


  SlblConfigFileName = '�ϥΪ��]�w��:';
  SlblConfigCASAddr = '[CAS�D��]:%s  %s Source Id:%s  %s Source Id:%s';
  SlblEnableControlSend = '�ҥ�[���O�ǰe]';
  SlblEnableFeedbackRecv = '�ҥ�[�^�Ǿ���]';

  SDockPanel1 = '�t�Υx���A����';
  SDockPanel2 = '���O�ǰe����';
  SDockPanel3 = '�R�O�Ҧ�����';
  SDockPanel4 = '�^�Ǿ������';
  SDockPanel5 = '�ҰʰT������';
  SDockPanel6 = '�t�Υx�T������';
  SDockPanel7 = '�ǰe���O�T������';
  SDockPanel8 = '�^�Ǿ���T������';

  SControlSendTreeColumn1 = '�t�Υx';
  SControlSendTreeColumn2 = '�Ǹ�';
  SControlSendTreeColumn3 = '�d��';
  SControlSendTreeColumn4 = '���O';
  SControlSendTreeColumn5 = '�ǰe';
  SControlSendTreeColumn6 = '�^��';
  SControlSendTreeColumn7 = '�ȪA�H��';
  SControlSendTreeColumn8 = '���~�X';
  SControlSendTreeColumn9 = '���~�T��';
  SControlSendTreeColumn10 = '�B�z�ɶ�';

  SControlConsoleTreeColumn1 = '�ǰe�Ǹ�';
  SControlConsoleTreeColumn2 = '�^���Ǹ�';
  SControlConsoleTreeColumn3 = '���O�榡';
  SControlConsoleTreeColumn4 = '�^��';

  SFeedbackTreecColumn1 = '�Ǹ�';
  SFeedbackTreecColumn2 = '�d��';
  SFeedbackTreecColumn3 = '���O';
  SFeedbackTreecColumn4 = '�g�J';
  SFeedbackTreecColumn5 = '�^��';
  SFeedbackTreecColumn6 = '�B�z�ɶ�';
  SFeedbackTreecColumn7 = '����';

  SMenuItem1 = '�ﶵ';
  SMenuItem11 = '�˵�/�ק�]�w�ɰѼ�';
  
  SMenuItem2 = '�˵�';
  SMenuItem21 = '���m�����t�m';
  SMenuItem22 = '�t�Υx-����';
  SMenuItem23 = '���O�ǰe-����';
  SMenuItem24 = '�R�O�Ҧ�-����';
  SMenuItem25 = '�^�Ǿ���-����';
  SMenuItem26 = '�ҰʰT��-����';
  SMenuItem27 = '�t�Υx-�T������';
  SMenuItem28 = '�ǰe���O-�T������';
  SMenuItem29 = '�^�Ǿ���-�T������';

  SMenuItem3 = '����';

  SButtonOK = '�T�w';
  SButtonCancel = '����';
  SButtonSave = '�x�s';
  SButtonDbTest = '����';
  SButtonAdd = '�s�W';
  SButtonDelete = '�R��';

  SAck = 'ACK';
  SNack = 'NACK';
  SUnknow = 'N/A';

  SDlgConfimQuit = '�T�{�����{��?';

  SDbTestParamIsEmpty = '�t�Υx[%s],��Ʈw�s���Ѽƭȿ�Ȧ��~�C';
  SDbTestError = '�t�Υx[%s],��Ʈw�s�����տ��~, �T��:%s�C';
  SDbTestNow = '�t�Υx[%s],��Ʈw�s�����դ�.......�C';
  SDbTestOk = '�t�Υx[%s],��Ʈw�s�����է����C';

implementation

end.
