unit cbResStr;

interface

resourcestring

  SVendor = '開博科技';
  SAppTitle = 'SMS Command Gateway';
  SVersion = '版本: 5.3 ( Build 5.3.0 )';
  
  SUnknownError = '未知錯誤。';
  SLicenceNotRegister = '本軟體尚未註冊,若需合法使用請連絡%s。';
  SLicenceExpire = '本軟體使用期限已過,若需使用請連絡%s。';
  SLicenceImmediateExpire = '本軟體已被限制使用,請連絡%s。';
  SLicenceInvalid = '本軟體註冊資料不合法。';

  SInputFormTitle = '軟體註冊';
  SLicenceNote = '請提供下列軟體序號,並輸入註冊號碼,以取得本軟體合法使用權。';
  SLicenceNote2 = '輸入軟體序號可產生註冊號碼。';
  SLicenceKey = '軟體序號';
  SRegisterKey = '註冊序號';
  SInstallDate = '安裝日期';
  SExpireDate = '使用到期日';
  SMacList = 'MAC序號';


  SConfigSearchFile = '搜尋設定檔中。';
  SConfigSearchCompelete = '共找到%d個設定檔。';

  SConfigFileNotAssign = '尚未指定設定檔。';
  SConfigFileNotExists = '指定的設定檔%s不存在, 請檢查檔案名稱或是路徑是否正確。';
  SConfigFileChange = '切換至設定檔-%s。';

  SConfigOpenSuccess = '設定檔%s載入完成。';
  SConfigOpenError = '設定檔%s載入錯誤。原因:%s';

  SConfigSoReadSuccess = '系統台讀取完成,共載入%d個系統台。';
  SConfigSoRead = '讀取系統台設定[%s]。';
  SConfigSoReadError = '設定檔,讀取系統台錯誤。原因:%s';

  SConfigFeedbackDbRead = '讀取回傳機制回寫資料庫設定。';
  SConfigFeedbackDbSuccess = '回傳機制回寫資料庫設定讀取完成。';

  SConfigSoWriteError = '設定檔[系統台]寫入錯誤。原因:%s';
  SConfigSoWriteSuccess = '設定檔[系統台]寫入完成。';

  SConfigCommonReadError = '設定檔,讀取一般設定錯誤。原因:%s';
  SConfigCommonWriteError = '設定檔,[一般設定]寫入錯誤。原因:%s。';
  SConfigCommonReadSuccess = '設定檔[一般設定]讀取完成。';
  SConfigCommonWritSuccess = '設定檔[一般設定]寫入完成。';

  SConfigCommonProcessRecordCount = '讀取設定[每次處理資料筆數]=%d';
  SConfigCommonProcessIPPV = '讀取設定[處理IPPV資料]=%s';
  SConfigCommonProcessBatch = '讀取設定[處理批次資料]=%s';
  SConfigCommonDbMultiThread = '讀取設定[各系統台使用獨立執行緒]=%s';
  SConfigCommonDbRetryFrequence = '讀取設定[系統台斷線時重連秒數]=%d';
  SConfigCommonDbWriteErrorWhenSocketFail = '讀取設定[無法傳送指令時,將資料標示錯誤]=%s';
  SConfigCommonDbResendCount = '讀取設定[資料重送次數限制]=%d';
  SConfigCommonBatchOperator = '讀取設定[批次資料識別帳號]=%s';
  SConfigCommonCASRetryFrequence = '讀取設定[CAS斷線時重連秒數]=%d';
  SConfigCommonBusyTimeStart = '讀取設定[忙碌時段起]=%s';
  SConfigCommonBusyTimeEnd = '讀取設定[忙碌時段迄]=%s';
  SConfigCommonBusyTimeDbReadFrequence = '讀取設定[忙碌時段讀取資料秒數]=%s';
  SConfigCommonNormalTimeDbReadFrequence = '讀取設定[一般時段讀取資料秒數]=%s';
  SConfigCommonCASSendErrMax = '讀取設定[傳送指令時, 發生多少次錯誤自動斷線]=%d';
  SConfigCommonCASRecvErrMax = '讀取設定[接收指令時, 發生多少次錯誤自動斷線]=%d';
  SConfigCommonCASCommCheck = '讀取設定[每多少送一次 Communication Check]=%d';

  SConfigHighCmdReading = '讀取設定[高階指令對照表]=%s(%s)。';
  SConfigHighCmdReadError = '設定檔,讀取[高階指令對照表]錯誤。原因:%s';
  SConfigHighCmdReadIsEmpty = '設定檔,讀取[高階指令對照表]錯誤。原因:指令對照表內容為空。';
  SConfigHighCmdReadSuccess = '[高階指令對照表]讀取完成。';
  SConfigHighCmdWriteError = '設定檔[高階指令對照表]寫入錯誤。原因:%s。';
  SConfigHighCmdWriteSuccess = '設定檔[高階指令對照表]寫入完成。';

  SConfigLowCmdReading = '讀取設定[低階指令描述表]=%s(%s)。';
  SConfigLowCmdReadError = '設定檔,讀取[低階指令描述表]錯誤。原因:%s';
  SConfigLowCmdReadIsEmpty = '設定檔,讀取[低階指令描述表]錯誤。原因:指令對照表內容為空。';
  SConfigLowCmdReadSuccess = '[低階指令描述表]讀取完成。';
  SConfigLowCmdWriteError = '設定檔[低階指令描述表]寫入錯誤。原因:%s。';
  SConfigLowCmdWriteSuccess = '設定檔[低階指令描述表]寫入完成。';

  SConfigCmdErrorReading = '讀取設定[指令錯誤對照表]。';
  SConfigCmdErrorReadSuccess = '[指令錯誤對照表]讀取完成。';
  SConfigCmdErrorWriteSuccess = '設定檔[指令錯誤對照表]寫入完成。';
  SConfigCmdErrorWriteError = '設定檔[指令錯誤對照表]寫入錯誤。原因:%s。';


  SConfigCmdTransUseSimulator = '讀取設定[使用模擬器]=%s。';
  SConfigCmdTransEnableControlSend = '讀取設定[啟用傳送指令]=%s。';
  SConfigCmdTransEnableFeedbackRecv = '讀取設定[啟用回傳機制]=%s。';
  SConfigCmdTransSendCommandDelay = '讀取設定[傳送低階指令延遲秒數]=%f。';
  SConfigCmdTransReadSuccess = '指令傳送設定讀取完成。';
  SConfigCmdTransReadError = '設定檔,讀取指令傳送設定錯誤。原因:%s。';
  SCOnfigCmdTransWriteError = '設定檔[指令傳送]寫入錯誤。原因:%s。';
  SCOnfigCmdTransWriteSuccess = '設定檔[指令傳送]寫入完成。';

  SConfigCASReadingIP = '讀取設定CAS主機[IP]=%s。';
  SConfigCASReadingSPort = '讀取設定CAS主機[Control Port]=%d。';
  SConfigCASReadingRPort = '讀取設定CAS主機[Feedback Port]=%d。';

  SConfigCASReadingControlSourceId = '讀取設定CAS主機[Control Source Id]=%s。';
  SConfigCASReadingControlDestId = '讀取設定CAS主機[Control Dest Id]=%s。';
  SConfigCASReadingControlMopPPId = '讀取設定CAS主機[Control MopPP Id]=%s。';

  SConfigCASReadingFeedbackSourceId = '讀取設定CAS主機[Feedback Source Id]=%s。';
  SConfigCASReadingFeedbackDestId = '讀取設定CAS主機[Feedback Dest Id]=%s。';
  SConfigCASReadingFeedbackMopPPId = '讀取設定CAS主機[Feedback MopPP Id]=%s。';

  SConfigCASReadControlTransId = '讀取設定CAS主機[Feedback Trans Id]=%s。';
  SConfigCASReadFeedbackTransId = '讀取設定CAS主機[Feedback Trans Id]=%s。';
  SConfigCASReadMailId = '讀取設定CAS主機[Mail Id]=%s。';
  SConfigCASReadCmdMaxCounter = '讀取設定CAS主機[低階指令傳送計數器上限:]=%d。';

  SConfigCASReadSuccess = 'CAS主機設定讀取完成。';
  SConfigCASWriteError = '設定檔寫入[CAS主機]錯誤。原因:%s';
  SConfigCASWritSuccess = '設定檔[CAS主機]設定寫入完成。';

  SConfigNoneSet = '未設定';

  SConfigBuildSoTreeStart = '建立系統台....';
  SConfigBuildSoTree = '建立系統台[%s]。';
  SConfigBuildSoTreeEnd = '系統台建立完成。';

  SAutoExec = '已啟用自動執行,稍後程式將自行啟動。';
  SNoneAutoExec = '自動執行未啟用,請手動執行。';
  SUnknowExec = '未啟用%s及%s,設定檔設定錯誤。';
  SExecMode = '已設定啟用%s';

  SDbConnectStart = '各系統台資料庫連線中....';
  SDbConnectSuccess = '系統台[%s]資料庫連結完成。';
  SDbConnectError = '系統台[%s]資料庫連結失敗。原因:%s。%d秒後將自動重新連結。';
  SDbConnectEnd = '系統台資料庫連線完成。';

  SDbDisConnectStart = '各系統台資料庫離線中。';
  SDbDisConnectSuccess = '系統台[%s]資料庫離線完成。';
  SDbDisConnectError = '系統台[%s]資料庫離線失敗。原因:%s';
  

  SFeedbackDbConnectStart = '[回傳機制]資料庫連線中....';
  SFeedbackDbConnectSuccess = '[回傳機制]資料庫連結完成。';
  SFeedbackDbConnectError = '[回傳機制]資料庫連結失敗。原因:%s。%d秒後將自動重新連結。';
  SFeedbackDbbConnectEnd = '[回傳機制]資料庫連線完成。';

  SFeedbackDbDisConnectStart = '[回傳機制]資料庫離線中。';
  SFeedbackDbDisConnectSuccess = '[回傳機制]資料庫離線完成。';
  SFeedbackDbDisConnectError = '[回傳機制]資料庫離線失敗。原因:%s。';
  SFeedbackDbWriteDataError = '[回傳機制]回傳資料寫入失敗。原因:%s。';
  SFeedbackDbWriteDataDone = '[回傳機制]回傳資料寫入完成,共寫入%d筆。';
  SFeedbackDbWriteDataError2 = '[回傳機制]回傳資料寫入失敗,廢棄%d筆。';

  SFeedbackDbRertyConnect = '[回傳機制]資料庫重新連結。';
  SFeedbackDbHasProblem = '[回傳機制]資料庫寫入已失敗多次,%d秒後將自動重新建立資料庫連結。';

  SDbGetDataError = '系統台[%s]存取來源資料失敗。原因:%s';
  SDbHasProblem = '系統台[%s]資料庫存取已失敗多次,%d秒後將自動重新建立資料庫連結。';
  SDbRertyConnect = '系統台[%s]資料庫重新連結。';
  SDbAddToControlSendList = '系統台[%s]待處理指令%d筆,已寫入佇列等候傳送。';
  SDbScanSo = '系統台[%s],無待處理資料。';
  SDbCheckSendListTooMany = '待處理指令佇列已超出指令基準數[%d/%d],系統台資料輪詢將等候下次處理。';

  SDbWriteAlreadySendError = '系統台[%s]回寫指令已傳送失敗,指令:%s,序號:%s。原因:%s。';
  SDbWriteAlreadySendCount = '系統台[%s]回寫指令狀態[處理中],本次處理%d筆。';
  SDbWriteNotSendCount = '系統台[%s]回寫指令狀態[無法傳送],本次處理%d筆。';
  SDbWriteAckError = '系統台[%s]回寫指令回應失敗,指令:%s,序號:%s。原因:%s。';
  SDbWriteAckCount = '系統台[%s]回寫指令狀態[已處理],本次完成%d筆。';
  SDbWriteLogCount = '系統台[%s]指令記錄日誌寫入完成,本次日誌記錄%d筆數。';


  SControlSendManager = '傳送指令佇列';
  SControlSendDoneManager = '已傳送指令佇列';
  SControlRecvManager = '指令回應佇列';
  SControlSendLogManager = '指令日誌佇列';
  SCleanupCommandManagerError = '清除[%s]發生錯誤。原因:%s。';
  SCleanupCommandManagerSuccess =  '清除[%s]完成。';


  SThreadStart = '啟始%s執行緒。';
  SThreadStartError = '啟始%s執行緒失敗。原因:%s。';
  SThreadReady = '%s執行緒就序,待命中。';
  SThreadStop = '正在停止%s執行緒。';
  SThreadStopError = '停止%s執行緒失敗。原因:%s。';
  SThreadEnd = '%s執行緒停止已完成。';

  SThreadHasProblem = '%s執行緒與CAS主機連結失敗, %d秒後將自動重新建立連結。';
  SDbThreadRunDelay = '系統台資料庫延後連結, 等候CAS主機連結。';

  SControlSend = '[傳送指令]';
  SControlRecv = '[接收回應]';
  SSoDb = '[系統台資料庫]';
  SFeedbackDb = '[回傳機制資料庫]';
  SFeedbackSend = '[回傳回應]';
  SFeedbackRecv = '[回傳資料]';
  SGuiCleanup = '[清除畫面顯示指令]';
  SReleaseControl = '[釋放主機]';

  SRunStop = '%s已停止執行。';

  SSocketBeginConnection = '%s主機 ( %s : %d ) 連結中。';
  SSocketConnectionFailure = '%s主機連結失敗。原因:TCP/IP連線失敗。';
  SSocketConnectionReject = '%s主機連結失敗。原因:NAGRA拒絕連線。';
  SSocketConnectionPostpend = '%s主機連結失敗。原因:NAGRA系統繁忙。';
  SSocketConnectionError = '%s主機連結失敗。原因:%s。';
  SSocketConnectionSuccess = '%s主機連結完成。';
  SSocketThreadError = '主機%s執行緒失敗。原因:%s。';

  SControlSendEndConnection = '[傳送指令]與主機 ( %s : %d ) 離線中。';
  SControlSendEndConnectionSuccess = '[傳送指令]主機離線完成。';
  SControlSendEndConnectionError = '[傳送指令]與主機離線失敗。原因:%s。';

  SSendCommandError = '%s失敗, 高階指令:%s, 低階指令:%s。原因:%s。';
  SSendErrors = '%s已失敗%d次以上,系統將停止傳送指令,並終止主機連結。';
  SCommunicationCheckError = '主機 ( %s : %d ) 拒絕%s,系統將停止%s,並終止連結。';
  SWaitRetry = '%s%d秒後將自動重新建立連結。';
  SEnsureCandSendCommand = '%s確認中。';
  SEnsureOK = '%s確認完成,可以開始%s。';
  SEnsureError = '%s確認失敗,%s被拒絕,系統將停止傳送指令,並終止連結。';
  SControlSendRecord = '%s已處理等待傳送佇列指令,高階指令%d筆,低階指令%d筆';
  SControlSendCasMaxCounterWarning1 = '%s已達到低階指令傳送計數器上限,將自動降低指令傳送速率,已避免 CAS 主機負荷過量。';
  SFeedbackSendRecord = '%s已處理回應佇列指令,共計%d筆。';


  SlblConfigFileName = '使用的設定檔:';
  SlblConfigCASAddr = '[CAS主機]:%s  %s Source Id:%s  %s Source Id:%s';
  SlblEnableControlSend = '啟用[指令傳送]';
  SlblEnableFeedbackRecv = '啟用[回傳機制]';

  SDockPanel1 = '系統台狀態視窗';
  SDockPanel2 = '指令傳送視窗';
  SDockPanel3 = '命令模式視窗';
  SDockPanel4 = '回傳機制視窗';
  SDockPanel5 = '啟動訊息視窗';
  SDockPanel6 = '系統台訊息視窗';
  SDockPanel7 = '傳送指令訊息視窗';
  SDockPanel8 = '回傳機制訊息視窗';

  SControlSendTreeColumn1 = '系統台';
  SControlSendTreeColumn2 = '序號';
  SControlSendTreeColumn3 = '卡號';
  SControlSendTreeColumn4 = '指令';
  SControlSendTreeColumn5 = '傳送';
  SControlSendTreeColumn6 = '回應';
  SControlSendTreeColumn7 = '客服人員';
  SControlSendTreeColumn8 = '錯誤碼';
  SControlSendTreeColumn9 = '錯誤訊息';
  SControlSendTreeColumn10 = '處理時間';

  SControlConsoleTreeColumn1 = '傳送序號';
  SControlConsoleTreeColumn2 = '回應序號';
  SControlConsoleTreeColumn3 = '指令格式';
  SControlConsoleTreeColumn4 = '回應';

  SFeedbackTreecColumn1 = '序號';
  SFeedbackTreecColumn2 = '卡號';
  SFeedbackTreecColumn3 = '指令';
  SFeedbackTreecColumn4 = '寫入';
  SFeedbackTreecColumn5 = '回應';
  SFeedbackTreecColumn6 = '處理時間';
  SFeedbackTreecColumn7 = '機號';

  SMenuItem1 = '選項';
  SMenuItem11 = '檢視/修改設定檔參數';
  
  SMenuItem2 = '檢視';
  SMenuItem21 = '重置版面配置';
  SMenuItem22 = '系統台-視窗';
  SMenuItem23 = '指令傳送-視窗';
  SMenuItem24 = '命令模式-視窗';
  SMenuItem25 = '回傳機制-視窗';
  SMenuItem26 = '啟動訊息-視窗';
  SMenuItem27 = '系統台-訊息視窗';
  SMenuItem28 = '傳送指令-訊息視窗';
  SMenuItem29 = '回傳機制-訊息視窗';

  SMenuItem3 = '關於';

  SButtonOK = '確定';
  SButtonCancel = '取消';
  SButtonSave = '儲存';
  SButtonDbTest = '測試';
  SButtonAdd = '新增';
  SButtonDelete = '刪除';

  SAck = 'ACK';
  SNack = 'NACK';
  SUnknow = 'N/A';

  SDlgConfimQuit = '確認結束程式?';

  SDbTestParamIsEmpty = '系統台[%s],資料庫連結參數值輸值有誤。';
  SDbTestError = '系統台[%s],資料庫連結測試錯誤, 訊息:%s。';
  SDbTestNow = '系統台[%s],資料庫連結測試中.......。';
  SDbTestOk = '系統台[%s],資料庫連結測試完成。';

implementation

end.
