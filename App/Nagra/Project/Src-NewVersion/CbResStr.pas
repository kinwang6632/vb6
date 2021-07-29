unit cbResStr;

interface

resourcestring
  SVendor = '�}�լ��';
  SAppTitle = 'SMS Command Gateway 20005';
  SUnknownError = '�������~�C';
  SLicenceNotRegister = '���n��|�����U,�Y�ݦX�k�ϥνгs��%s�C';
  SLicenceExpire = '���n��ϥδ����w�L,�Y����νгs��%s�C';
  SLicenceImmediateExpire = '���n��w�Q����ϥ�,�гs��%s�C';
  SLicenceInvalid = '���n����U��Ƥ��X�k%s�C';

  SConfigNoneRead = '�|��Ū�����ҳ]�w��%s�C';
  SConfigFileNotExists = '���w�����ҳ]�w��%s���s�b�C';

  SConfigOpenSuccess = '���ҳ]�w��%s���J�����C';
  SConfigOpenError = '���ҳ]�w��%s���J���~�C��]:%s';

  SConfigSoReadSuccess = '�t�ΥxŪ������,�@���J%d�Өt�Υx�C';
  SConfigSoRead = 'Ū���t�Υx�]�w[%s]�C';
  SConfigSoReadError = '���ҳ]�w��,Ū���t�Υx���~�C��]:%s';

  SConfigSoWriteError = '���ҳ]�w��,�g�J�t�Υx���~�C��]:%s';
  SConfigSoWrite = '�g�J�t�Υx�]�w[%s]�C';
  SConfigSoWriteSuccess = '�t�Υx�g�J����,�@�g�J%d�Өt�Υx�C';

  SConfigCommonReadError = '���ҳ]�w��,Ū���@��]�w���~�C��]:%s';
  SConfigCommonWriteError = '���ҳ]�w��,�g�J�@��]�w���~�C��]:%s';
  SConfigCommonReadSuccess = '�@��]�wŪ�������C';
  SConfigCommonWritSuccess = '�@��]�w�g�J�����C';

  SConfigCommonProcessRecordCount = 'Ū���]�w[�C���B�z��Ƶ���]=%d';
  SConfigCommonProcessIPPV = 'Ū���]�w[�B�zIPPV���]=%s';
  SConfigCommonProcessBatch = 'Ū���]�w[�B�z�妸���]=%s';
  SConfigCommonDbMultiThread = 'Ū���]�w[�U�t�Υx�ϥοW�߰����]=%s';
  SConfigCommonDbRetryFrequence = 'Ū���]�w[�t�Υx�_�u�ɭ��s���]=%d';
  SConfigCommonDbResendCount = 'Ū���]�w[��ƭ��e���ƭ���]=%d';
  SConfigCommonBatchOperator = 'Ū���]�w[�妸����ѧO�b��]=%s';
  SConfigCommonBusyTimeStart = 'Ū���]�w[���L�ɬq�_]=%s';
  SConfigCommonBusyTimeEnd = 'Ū���]�w[���L�ɬq��]=%s';
  SConfigCommonBusyTimeDbReadFrequence = 'Ū���]�w[���L�ɬqŪ����Ƭ��]=%s';
  SConfigCommonNormalTimeDbReadFrequence = 'Ū���]�w[�@��ɬqŪ����Ƭ��]=%s';

  SConfigCmdReading = 'Ū���]�w[���O��Ӫ�]=%s(%s)�C';
  SConfigCmdReadError = '���ҳ]�w��,Ū�����O��Ӫ���~�C��]:%s';
  SConfigCmdReadIsEmpty = '���ҳ]�w��,Ū�����O��Ӫ���~�C��]:���O��Ӫ��e���šC';
  SConfigCmdReadSuccess = '���O��Ӫ�Ū�������C';
  SConfigCmdWriteError = '���ҳ]�w��,�g�J���O��Ӫ���~�C��]:%s';
  SConfigCmdWriteSuccess = '���O��Ӫ�g�J�����C';

  SConfigCASReadingIP = 'Ū���]�wNAGRA�D��[IP]=%s�C';
  SConfigCASReadingSPort = 'Ū���]�wNAGRA�D��[Control Port]=%d�C';
  SConfigCASReadingRPort = 'Ū���]�wNAGRA�D��[Feedback Port]=%d�C';
  SConfigCASReadingSourceId = 'Ū���]�wNAGRA�D��[Source Id]=%s�C';
  SConfigCASReadingDestId = 'Ū���]�wNAGRA�D��[Dest Id]=%s�C';
  SConfigCASReadingMopPPId = 'Ū���]�wNAGRA�D��[MopPP Id]=%s�C';
  SConfigCASReadSuccess = 'NAGRA�D���]�wŪ�������C';

  SConfigNoneSet = '���]�w';


  SConfigBuildSoTreeStart = '�إߨt�Υx....';
  SConfigBuildSoTree = '�إߨt�Υx[%s]�C';
  SConfigBuildSoTreeEnd = '�إߨt�Υx�����C';

  SAutoExec = '�w�ҥΦ۰ʰ���,�y��{���N�ۦ�ҰʡC';
  SNoneAutoExec = '�۰ʰ��楼�ҥ�,�Ф�ʰ���C';

  SDbConnectStart = '�U�t�Υx��Ʈw�s�u��....';
  SDbConnectSuccess = '�t�Υx[%s]��Ʈw�s�������C';
  SDbConnectError = '�t�Υx[%s]��Ʈw�s�����ѡC��]:%s';
  SDbConnectEnd = '�t�Υx��Ʈw�s�u�����C';

  SDbDisConnectStart = '�U�t�Υx��Ʈw���u��....';
  SDbDisConnectSuccess = '�t�Υx[%s]��Ʈw���u�����C';
  SDbDisConnectError = '�t�Υx[%s]��Ʈw���u���ѡC��]:%s';
  SDbDisConnectThreadError = '�פ�t�Υx��Ʈw��������ѡC��]:%s';
  SDbDisConnectEnd = '�t�Υx��Ʈw���u�����C';


  SDbGetDataError = '�t�Υx[%s]�s���ӷ���ƥ��ѡC��]:%s';
  SDbHasProblem = '�t�Υx[%s]��Ʈw�s���w���Ѧh��,%d���N�۰ʭ��s�إ߸�Ʈw�s���C';
  SDbRertyConnect = '�t�Υx[%s]��Ʈw���s�s���C';
  SDbAddToControlSendList = '�t�Υx[%s]�ݳB�z���O%d��,�w�g�J��C���Զǰe�C';

  SDbWriteAlreadySendError = '�t�Υx[%s]�^����O�w�ǰe����,���O:%s,�Ǹ�:%s�C��]:%s�C';
  SDbWriteAlreadySendCount = '�t�Υx[%s]�^����O���A[�B�z��],��������%d���C';
  SDbWriteAckError = '�t�Υx[%s]�^����O�^������,���O:%s,�Ǹ�:%s�C��]:%s�C';
  SDbWriteAckCount = '�t�Υx[%s]�^����O���A[�w�B�z],��������%d���C';
  SDbWriteLogCount = '�t�Υx[%s]���O�O����x�g�J����,������x�O��%d���ơC';

  SControlSendManager = '�ǰe���O��C';
  SControlSendDoneManager = '�w�ǰe���O��C';
  SControlRecvManager = '���O�^����C';
  SControlSendLogManager = '���O��x��C';
  SCleanupCommandManagerError = '�M��[%s]�o�Ϳ��~�C��]:%s�C';
  SCleanupCommandManagerSuccess =  '�M��[%s]�����C';


  SControlSendOpen = '�ҩl�ǰe���O������C';
  SControlSendOpenReady = '�ǰe���O������N��,�ݩR���C';
  SControlSendClose = '���b����ǰe���O������C';
  SControlSendCloseError = '����ǰe���O��������ѡC��]:%s�C';
  SControlSendCloseDone = '�ǰe���O���������w�����C';


  SControlSendBeginConnection = '�ǰe���O�D��[%s:%d]�s�����C';
  SControlSendConnectionReject = '�ǰe���O�D���s�����ѡC��]:NAGRA�ڵ��s�u�C';
  SControlSendConnectionPostpend = '�ǰe���O�D���s�����ѡC��]:NAGRA�t���c���C';
  SControlSendConnectionError = '�ǰe���O�D���s�����ѡC��]:%s�C';
  SControlSendConnectionSuccess = '�ǰe���O�D���s�������C';

  SControlSendEndConnection = '�ǰe���O�P�D��[%s:%d]���u���C';
  SControlSendEndConnectionSuccess = '�ǰe���O�P�D�����u�����C';
  SControlSendEndConnectionError = '�ǰe���O�P�D�����u���ѡC��]:%s�C';


  SlblConfigFileName = '�ϥγ]�w��:%s';

  SDockPanel1 = '�t�Υx���A';
  SDockPanel2 = '';
  SDockPanel3 = '���O�Ҧ�����';
  SDockPanel4 = '�R�O�Ҧ�����';
  SDockPanel5 = '�ҰʰT������';
  SDockPanel6 = '�t�Υx�T������';
  SDockPanel7 = '�ǰe���O�T������';

  SAppStartup = '�{���w�ҰʡC';
  SAppRunning = '�{�����椤�C';
  SAppCloseing = '�{���������C';

  SDlgConfimQuit = '�T�{�����{��?';

implementation

end.
