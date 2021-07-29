unit cbResStr;

interface

resourcestring
  SVendor = '開博科技';
  SAppTitle = 'SMS Command Gateway 20005';
  SUnknownError = '未知錯誤。';
  SLicenceNotRegister = '本軟體尚未註冊,若需合法使用請連絡%s。';
  SLicenceExpire = '本軟體使用期限已過,若需續用請連絡%s。';
  SLicenceImmediateExpire = '本軟體已被限制使用,請連絡%s。';
  SLicenceInvalid = '本軟體註冊資料不合法%s。';

  SConfigNoneRead = '尚未讀取環境設定檔%s。';
  SConfigFileNotExists = '指定的環境設定檔%s不存在。';

  SConfigOpenSuccess = '環境設定檔%s載入完成。';
  SConfigOpenError = '環境設定檔%s載入錯誤。原因:%s';

  SConfigSoReadSuccess = '系統台讀取完成,共載入%d個系統台。';
  SConfigSoRead = '讀取系統台設定[%s]。';
  SConfigSoReadError = '環境設定檔,讀取系統台錯誤。原因:%s';

  SConfigSoWriteError = '環境設定檔,寫入系統台錯誤。原因:%s';
  SConfigSoWrite = '寫入系統台設定[%s]。';
  SConfigSoWriteSuccess = '系統台寫入完成,共寫入%d個系統台。';

  SConfigCommonReadError = '環境設定檔,讀取一般設定錯誤。原因:%s';
  SConfigCommonWriteError = '環境設定檔,寫入一般設定錯誤。原因:%s';
  SConfigCommonReadSuccess = '一般設定讀取完成。';
  SConfigCommonWritSuccess = '一般設定寫入完成。';

  SConfigCommonProcessRecordCount = '讀取設定[每次處理資料筆數]=%d';
  SConfigCommonProcessIPPV = '讀取設定[處理IPPV資料]=%s';
  SConfigCommonProcessBatch = '讀取設定[處理批次資料]=%s';
  SConfigCommonDbMultiThread = '讀取設定[各系統台使用獨立執行緒]=%s';
  SConfigCommonDbRetryFrequence = '讀取設定[系統台斷線時重連秒數]=%d';
  SConfigCommonDbResendCount = '讀取設定[資料重送次數限制]=%d';
  SConfigCommonBatchOperator = '讀取設定[批次資料識別帳號]=%s';
  SConfigCommonBusyTimeStart = '讀取設定[忙碌時段起]=%s';
  SConfigCommonBusyTimeEnd = '讀取設定[忙碌時段迄]=%s';
  SConfigCommonBusyTimeDbReadFrequence = '讀取設定[忙碌時段讀取資料秒數]=%s';
  SConfigCommonNormalTimeDbReadFrequence = '讀取設定[一般時段讀取資料秒數]=%s';

  SConfigCmdReading = '讀取設定[指令對照表]=%s(%s)。';
  SConfigCmdReadError = '環境設定檔,讀取指令對照表錯誤。原因:%s';
  SConfigCmdReadIsEmpty = '環境設定檔,讀取指令對照表錯誤。原因:指令對照表內容為空。';
  SConfigCmdReadSuccess = '指令對照表讀取完成。';
  SConfigCmdWriteError = '環境設定檔,寫入指令對照表錯誤。原因:%s';
  SConfigCmdWriteSuccess = '指令對照表寫入完成。';

  SConfigCASReadingIP = '讀取設定NAGRA主機[IP]=%s。';
  SConfigCASReadingSPort = '讀取設定NAGRA主機[Control Port]=%d。';
  SConfigCASReadingRPort = '讀取設定NAGRA主機[Feedback Port]=%d。';
  SConfigCASReadingSourceId = '讀取設定NAGRA主機[Source Id]=%s。';
  SConfigCASReadingDestId = '讀取設定NAGRA主機[Dest Id]=%s。';
  SConfigCASReadingMopPPId = '讀取設定NAGRA主機[MopPP Id]=%s。';
  SConfigCASReadSuccess = 'NAGRA主機設定讀取完成。';

  SConfigNoneSet = '未設定';


  SConfigBuildSoTreeStart = '建立系統台....';
  SConfigBuildSoTree = '建立系統台[%s]。';
  SConfigBuildSoTreeEnd = '建立系統台完成。';

  SAutoExec = '已啟用自動執行,稍後程式將自行啟動。';
  SNoneAutoExec = '自動執行未啟用,請手動執行。';

  SDbConnectStart = '各系統台資料庫連線中....';
  SDbConnectSuccess = '系統台[%s]資料庫連結完成。';
  SDbConnectError = '系統台[%s]資料庫連結失敗。原因:%s';
  SDbConnectEnd = '系統台資料庫連線完成。';

  SDbDisConnectStart = '各系統台資料庫離線中....';
  SDbDisConnectSuccess = '系統台[%s]資料庫離線完成。';
  SDbDisConnectError = '系統台[%s]資料庫離線失敗。原因:%s';
  SDbDisConnectThreadError = '終止系統台資料庫執行緒失敗。原因:%s';
  SDbDisConnectEnd = '系統台資料庫離線完成。';


  SDbGetDataError = '系統台[%s]存取來源資料失敗。原因:%s';
  SDbHasProblem = '系統台[%s]資料庫存取已失敗多次,%d秒後將自動重新建立資料庫連結。';
  SDbRertyConnect = '系統台[%s]資料庫重新連結。';
  SDbAddToControlSendList = '系統台[%s]待處理指令%d筆,已寫入佇列等候傳送。';

  SDbWriteAlreadySendError = '系統台[%s]回填指令已傳送失敗,指令:%s,序號:%s。原因:%s。';
  SDbWriteAlreadySendCount = '系統台[%s]回填指令狀態[處理中],本次完成%d筆。';
  SDbWriteAckError = '系統台[%s]回填指令回應失敗,指令:%s,序號:%s。原因:%s。';
  SDbWriteAckCount = '系統台[%s]回填指令狀態[已處理],本次完成%d筆。';
  SDbWriteLogCount = '系統台[%s]指令記錄日誌寫入完成,本次日誌記錄%d筆數。';

  SControlSendManager = '傳送指令佇列';
  SControlSendDoneManager = '已傳送指令佇列';
  SControlRecvManager = '指令回應佇列';
  SControlSendLogManager = '指令日誌佇列';
  SCleanupCommandManagerError = '清除[%s]發生錯誤。原因:%s。';
  SCleanupCommandManagerSuccess =  '清除[%s]完成。';


  SControlSendOpen = '啟始傳送指令執行緒。';
  SControlSendOpenReady = '傳送指令執行緒就序,待命中。';
  SControlSendClose = '正在停止傳送指令執行緒。';
  SControlSendCloseError = '停止傳送指令執行緒失敗。原因:%s。';
  SControlSendCloseDone = '傳送指令執行緒停止已完成。';


  SControlSendBeginConnection = '傳送指令主機[%s:%d]連結中。';
  SControlSendConnectionReject = '傳送指令主機連結失敗。原因:NAGRA拒絕連線。';
  SControlSendConnectionPostpend = '傳送指令主機連結失敗。原因:NAGRA系統繁忙。';
  SControlSendConnectionError = '傳送指令主機連結失敗。原因:%s。';
  SControlSendConnectionSuccess = '傳送指令主機連結完成。';

  SControlSendEndConnection = '傳送指令與主機[%s:%d]離線中。';
  SControlSendEndConnectionSuccess = '傳送指令與主機離線完成。';
  SControlSendEndConnectionError = '傳送指令與主機離線失敗。原因:%s。';


  SlblConfigFileName = '使用設定檔:%s';

  SDockPanel1 = '系統台狀態';
  SDockPanel2 = '';
  SDockPanel3 = '指令模式視窗';
  SDockPanel4 = '命令模式視窗';
  SDockPanel5 = '啟動訊息視窗';
  SDockPanel6 = '系統台訊息視窗';
  SDockPanel7 = '傳送指令訊息視窗';

  SAppStartup = '程式已啟動。';
  SAppRunning = '程式執行中。';
  SAppCloseing = '程式結束中。';

  SDlgConfimQuit = '確認結束程式?';

implementation

end.
