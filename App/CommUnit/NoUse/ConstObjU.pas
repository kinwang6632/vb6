unit ConstObjU;

interface

type
    TIrdInfoObj = class(TObject)
      s_IrdCmdID : String;
      s_IrdOperation :String;
      s_IrdDataLength :String;
      s_IrdData :String;

    end;


    TReturnPathDataObj = class(TObject)
      Data : String;
    end;

    TSeqNoAndTransInfoObj = class(TObject) //new project
      SeqNo : String;
    end;

    TCompAndIpInfoObj = class(TObject) //new project
      IP : String;
    end;

    TNetworkIDInfoObj = class(TObject)
      NetworkID : String;
      Operation : String;  
    end;

    
    TCmdInfoObject = class(TObject)
        HIGH_LEVEL_CMD_ID           : String;
        ICC_NO                      : String;
        STB_NO                      : String;
        SUBSCRIPTION_BEGIN_DATE     : String;
        SUBSCRIPTION_END_DATE       : String;
        CMD_STATUS                  : String;
        OPERATOR                    : String;
        UPDTIME                     : String;
        ERR_CODE                    : String;
        ERR_MSG                     : String;
        IMS_PRODUCT_ID              : String;
        ZIP_CODE                    : String;
        CREDIT_MODE                 : String;
        CREDIT                      : String;
        CREDIT_LIMIT                : String;
        THRESHOLD_CREDIT            : String;
        EVENT_NAME                  : String;
        PRICE                       : String;
        CC_NUMBER_1                 : String;
        IP_ADDR                     : String;
        CC_PORT                     : String;
        CALLBACK_DATE               : String;
        CALLBACK_TIME               : String;
        CALLBACK_FREQUENCY          : String;
        FIRST_CALLBACK_DATE         : String;
        PHONE_NUM_1                 : String;
        PHONE_NUM_2                 : String;
        PHONE_NUM_3                 : String;
        SEQNO                       : String;
        MIS_IRD_CMD_ID              : String;
        MIS_IRD_CMD_DATA            : String;
        COMPCODE                    : String;
        RESENT_TIMES                : String;
        PROCESSING_DATE             : String;
        NOTES                       : String;
        COMPNAME                    : String;
        CLEANUP_DATE                : String;
        CONDITION_DATE              : String;
        COLLECT_DATE                : String;
        STB_FLAG                    : String;       
    end;

  TTempRecord = record
 	TempAry: array[0..5] of char;
  end;

  
  TDeviceData = record
//    len : Smallint;
    len : array[0..1] of char;
    data : array[0..1024] of char;
  end;

(*
  TDeviceCall = record
    len : array[0..1] of Char;
    op_mode : Char;
    obj_name_len : Char;
    obj_name : array[0..31] of Char;
    user_data: array [0..254] of Char;
  end;

  TDeviceAnswer = record
    len : Byte;
    code: Byte;
    data: array [0..1024] of Char;
  end;
*)

  TDeviceStatus = record
    len : Smallint; //signed 16-bit
    status : char; //8 bits
  end;



  TDeviceCall = record
    len : array[1..2] of Byte;
    op_mode : Byte;
    obj_name_len : Byte;
    obj_name : array[1..32] of Char;
  end;


  TDeviceAnswer = record
    statuslen : array [1..2] of Byte;
    status: Byte;
    answerlen: array [1..2] of Byte;
    code : Byte;
  end;


const
//以下是 Ca Dispatcher 傳給 CaProcessor 的參數長度變數
    HIGH_LEVEL_CMD_ID_LENGTH = 4;
    ICC_NO_LENGTH = 12;
    STB_NO_LENGTH = 16;
    SUB_BEGIN_DATE_LENGTH = 8;
    SUB_END_DATE_LENGTH = 8;
    CMD_STATUS_LENGTH = 1;
    OPERATOR_LENGTH = 20;
    IMS_PRODUCTID_LENGTH = 12;
    ZIP_CODE_LENGTH = 5;
    CREDIT_MODE_LENGTH = 1;
    CREDIT_LENGTH = 7;
    CREDIT_LIMIT_LENGTH = 7;
    THRESHOLD_CREDIT_LENGTH = 7;
    EVENT_NAME_LENGTH = 30;
    PRICE_LENGTH = 5;
    CC_NUMBER_LENGTH = 16;
    IP_LENGTH = 15;
    PORT_LENGTH = 5;
    CALLBACK_DATE_LENGTH = 8;
    CALLBACK_TIME_LENGTH = 6;
    CALLBACK_FREQUENCY_LENGTH = 2;
    FIRST_CALLBACK_DATE_LENGTH = 8;
    VERIFY_PHONE_NUMBER1_LENGTH = 16;
    VERIFY_PHONE_NUMBER2_LENGTH = 16;
    VERIFY_PHONE_NUMBER3_LENGTH = 16;
    MIS_IRD_CMD_ID_LENGTH = 2;
    SEQ_NO_LENGTH = 8;
    COMP_CODE_LENGTH = 3;
    RESENT_TIMES_LENGTH = 3;
    PROCESSING_DATE_LENGTH = 14;
    NOTES_LENGTH = 4;
    MIS_IRD_CMD_DATA_LENGTH = 3;
    COMP_NAME_LENGTH = 2;

//以上是 Ca Dispatcher 傳給 CaProcessor 的參數長度變數

    SEND_CMD_ERROR_FLAG = -999; //一定要是負值
    RE_CONNECT_NAGRA_ERROR_FLAG = -998; //一定要是負值
    CLOSE_SEND_NAGRA_FLAG = -997; //一定要是負值
    EXECUTE_SEND_NAGRA_FLAG = -996; //一定要是負值
    EXECUTE_ORACLE_SP_ERROR_FLAG = -995; //一定要是負值
    RECEIVE_CMD_RESPONSE_ERROR_FLAG = -994; //一定要是負值
    HANDLE_RESPONSE_ERROR_FLAG = -993; //一定要是負值
    WRITE_TO_MIS_ERROR_FLAG = -992; //一定要是負值
    UPDATE_CMD_STATUS_ERROR_FLAG = -989; //一定要是負值
    SOCKET_NOT_READY_FOR_READING_FLAG = -987; //一定要是負值
    REBUILD_SEND_CMD_CONNECTION_TO_NAGRA = -986; //一定要是負值
    REBUILD_NAGRA_CONNECTION_FAILED = -985; //一定要是負值    



    RECEIVED_DATA_SEP_FLAG = '^';
    ENCRYPTION_SEP = 'H';
    ENCRYPTION_KEY = '1234';
    REGISTRY_ROOT = 'SOFTWARE\Windows\Sys';
    SYS4_Y = 'A';
    SYS4_N = '5';
    SYS5_Y = '9';
    SYS5_N = 'B';

    SEND_TO_DISPATCHER_INFO_DEFAULT_TYPE = 0;
    SEND_TO_DISPATCHER_INFO_TYPE_1 = 1; //指令處理結果之 flag
    SEND_TO_DISPATCHER_INFO_TYPE_2 = 2; //指令之 log flag
    SEND_TO_DISPATCHER_INFO_TYPE_3 = 3; //response 之 log flag    
    
    SOCKET_WAIT_FOR_DATA_TIMEOUT = 2 ;// WaitForData 的 timeout 秒數 => 真正的值(WaitForTimeout)應該是此參數值再加上 frmMain.nG_CmdRefreshRate1 或 frmMain.nG_CmdRefreshRate2 (由尖峰或離峰時間而定)
    BATCH_OPERATOR = 'Batch Operation';
    DELETE_MAPPING_DATA_TIME = '0200'; // 表示凌晨 02:00 時,程式會將 CaSeqNoTransNumMappingData 中 UPDTIME=昨天的 data 刪除掉
    MAX_COMMAND_COUNT_PER_TIMES = 100;
    
    DB_LOG_FILE_NAME = 'CaLog.txt';
    MAX_SHOW_MEMO_ERROR_COUNT=1000;
    LOW_CREDIT_ALARM_CMD_ID = 200;
    CURRENT_DEBIT_AND_CREDIT_CMD_ID = 201;
    PPV_PURCHASE_LIST_CMD_ID = 202;
    PHONE_DISCREPANCIES_CMD_ID = 205;
    STB_RESPONDING_STATUS_CMD_ID = 206;
    ICC_MEMORY_FULL_ALARM_CMD_ID = 207;
    EVENT_DEFINITION_ERROR_CMD_ID = 208;
    NULL_EVENT_ERROR_CMD_ID = 209;
    EPG_DATA_FEED_FORMAT_ERROR_CMD_ID = 210;
    START_OF_REPORT_CMD_ID = 211;
    END_OF_REPORT_CMD_ID = 212;
    EVENT_PRODUCT_SCHEDULE_CHANGE_CMD_ID = 213;
    EVENT_REMOVE_ERROR_CMD_ID = 214;
    PRODUCTS_LIST_CMD_ID = 215;

    MAX_MESSAGE_LENGTH = 45;

    MAX_MESSAGE_LENGTH_NEW = 42;

    CA_UI_MAX_COUNT=500;
    CA_TESTING_CMD_ID='Z1';
    TMP_TIME_STAMP_FILE_NAME = 'ts.txt';    
    TIME_STAMP_FILE_NAME = 'TS.dat';
    INVOKE_MONITOR_OK_CODE = '0';
    THREAD_INFO_FILE = 'ThreadInfo.txt';
    SOCKET_INFO_FILE = 'SocketInfo.txt';
    NETWORK_OPERATION_194 = '194';
    NETWORK_OPERATION_195 = '195';
    NETWORK_OPERATION_198 = '198';
        
    SEND_NAGRA_DB_NAME = 'dbSendNagra';
    CABLE_NAGRA_INI_FILE_NAME = 'CableNagra.ini';
    LANGUAGE_SETTING_INI_FILE_NAME= 'LanguageSetting.ini';
    
    STR_SEP = '''';
    CMD_FIELD_SEP = '-';

    SUB_ITEM_COUNT = 9;

    
    SUPER_USER_ID='sys';
    SUPER_USER_PASSWD='sys';

    SOURCE_ID_LENGTH = 4;
    DEST_ID_LENGTH = 4;
    MOP_PPID_LENGTH = 5;

    ACKNOWLEDGE_CMD_ID = 1000;
    NON_ACKNOWLEDGE_CMD_ID = 1001;

    NO_ERROR = '0';
    NO_ERROR_MSG = '';
    COMP_ID_SEP = ','; // new project

    BUILD_CONNECTION_ERROR = 'E011';
    BUILD_CONNECTION_ERROR_MSG = 'build tcp connection error...';

    SEND_COMMAND_ERROR = 'E012';
    SEND_COMMAND_ERROR_MSG = 'CA command gateway already got the command,but could send to  Nagra system...';

    WRITE_TO_MIS_ERROR = 'E013';
    WRITE_TO_MIS_ERROR_MSG = 'Could not write command to mis...';

    NOTES_FIELD_FORMAT_ERROR = 'E014';
    NOTES_FIELD_FORMAT_ERROR_MSG = 'Send_Nagra.Notes format error...';

    STB_NO_LENGTH_ERROR = 'E015';
    STB_NO_LENGTH_ERROR_MSG = 'STB No length error...';

    ICC_CARD_NO_LENGTH_ERROR = 'E016';
    ICC_CARD_NO_LENGTH_ERROR_MSG = 'ICC Card No length error...';

    IMS_PRODUCT_ID_LENGTH_ERROR = 'E017';
    IMS_PRODUCT_ID_LENGTH_ERROR_MSG = 'IMS Product ID length error...';

    TRANS_IRD_ERROR = 'E021';
    TRANS_IRD_ERROR_MSG = 'IRD command process error...';

    OTHER_ERROR = 'E099';
    OTHER_ERROR_MSG = 'other error...';

    LISTVIEW_SUB_ITEM_COUNT = 10;
    IRD_COMMAND_DATA_LENGTH = 96;
    COMMAND_ID_LENGTH = 4;
    IMS_PRODUCT_ID_LENGTH = 12;
    BOXNO_ID_LENGTH = 14;
    ICC_CARDNO_LENGTH = 10;

    CALL_COLLECTOR_PHONE_NUM_LENGTH = 16;
    AUTHORIZED_PHONE_NUM_LENGTH = 16;
    EMM_UNIQUE_ADDRESS_TYPE = 'U';
    SEND_TRANSACTION_NUM_LENGTH = 9;
    TESTING_CMD_COMP_CODE = '099';
    TESTING_CMD_COMP_CODE2 = '100';
    FEEDBACK_COMP_CODE = '101';
    //PORT_60003_COMP_CODE = '98';
    SEQ_NUM_LENGTH = 6;
    PASSWORD_CHECKING_TIMES = 3;
    PRODUCT_INFO_FILENAME = 'product.dat';
    FEEDBACK_INFO_FILENAME = 'feedback.dat';
    OPERATION_INFO_FILENAME = 'operation.dat';
    CONTROL_INFO_FILENAME = 'control.dat';
    EMM_INFO_FILENAME = 'emm.dat';
    SYS_PARAM_FILENAME = 'Param.dat';
{
    PROCESSOR_SYS_PARAM_FILENAME = 'ProcessorParam.dat';
    DISPATCHER_SYS_PARAM_FILENAME = 'DispatcherParam.dat';
}
    USER_INFO_FILENAME = 'user.dat';
    NETWORK_ID_INFO_FILENAME = 'networkid.dat';
    
    EMM_CommandBeginID = 2;
    EMM_CommandEndID = 97;

    CONTROL_CommandBeginID = 100;
    CONTROL_CommandEndID = 111;

    FEEDBACK_CommandBeginID = 200;
    FEEDBACK_CommandEndID = 215;

    PRODUCT_CommandBeginID = 300;
    PRODUCT_CommandEndID = 309;

    OPERATION_CommandBeginID = 1000;
    OPERATION_CommandEndID = 1002;

implementation

end.
