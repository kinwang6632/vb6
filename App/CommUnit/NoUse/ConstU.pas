unit ConstU;

interface

type
    TCmdInfoObject = class(TObject)
        SeqNo                       : String;
        CompCode                    : String;
        CompName                    : String;
        HighLLevelCmdID             : String;
        SubscriberID                : String;
        IccNo                       : String;
        SubscriptionBeginDate       : String;
        SubscriptionEndDate         : String;
        PinCode                     : String;
        PopulationID                : String;
        RegionKey                   : String;
        ReportbackAvailability      : String;
        ReportbackDate              : String;
        Notes                       : String;
        Operator                    : String;
        PinControl                  : String;
        ServiceID                   : String;
    end;
    TCompAndIpInfoObj = class(TObject) //new project
      IP : String;
    end;

    TSignatureKeyInfoObj = class(TObject)
      sSignatureKey:String;
    end;

    TSequenceIdMappingInfoObj = class(TObject)
      CompCode : String;
      HighLevelCmdId : String;
      DispatcherSequenceNo : String;
    end;

    TErrorMsgObj = class(TObject)
      ErrorMsg : String;
    end;
const
//琌 Ca Dispatcher 肚倒 CaProcessor 把计跑计
    SUBSCRIBER_ID_LENGTH = 8;
    HIGH_LEVEL_CMD_ID_LENGTH = 4;
    REGION_KEY_LENGTH = 8;
    REPORTBACK_AVAILABILITY_LENGTH = 1;
    ICC_NO_LENGTH = 12;
//    STB_NO_LENGTH = 16;
    PIN_CODE_LENGTH = 4;
    POPULATION_ID_LENGTH = 2;
    REPORTBACK_DATE_LENGTH = 2;

    SUB_BEGIN_DATE_LENGTH = 8;
    SUB_END_DATE_LENGTH = 8;
    CMD_STATUS_LENGTH = 1;
    OPERATOR_LENGTH = 10;

    SEQ_NO_LENGTH = 8;
    COMP_CODE_LENGTH = 3;
    COMP_NAME_LENGTH = 10;
//    RESENT_TIMES_LENGTH = 3;
//    PROCESSING_DATE_LENGTH = 14;
    NOTES_LENGTH = 4;
    PIN_CONTROL_LENGTH = 1;
    SERVICE_ID_LENGTH = 5;
//    MIS_IRD_CMD_DATA_LENGTH = 3;

//琌 Ca Dispatcher 肚倒 CaProcessor 把计跑计



    SEND_CMD_ERROR_FLAG = -999; //﹚璶琌璽
    RE_CONNECT_NDS_ERROR_FLAG = -998; //﹚璶琌璽
    CLOSE_SEND_NDS_FLAG = -997; //﹚璶琌璽
    EXECUTE_SEND_NDS_FLAG = -996; //﹚璶琌璽
    EXECUTE_ORACLE_SP_ERROR_FLAG = -995; //﹚璶琌璽
    RECEIVE_CMD_RESPONSE_ERROR_FLAG = -994; //﹚璶琌璽
    HANDLE_RESPONSE_ERROR_FLAG = -993; //﹚璶琌璽
    WRITE_TO_MIS_ERROR_FLAG = -992; //﹚璶琌璽
    UPDATE_CMD_STATUS_ERROR_FLAG = -989; //﹚璶琌璽
    SOCKET_NOT_READY_FOR_READING_FLAG = -987; //﹚璶琌璽
    REBUILD_SEND_CMD_CONNECTION_TO_NDS = -986; //﹚璶琌璽
    REBUILD_NDS_CONNECTION_FAILED = -985; //﹚璶琌璽



    
    CURRENCY_LENGTH = 4;
    PERSONAL_REGION_LENGTH_LENGTH = 2;
    REPORTBACK_TIME_LENGTH = 6;    
    ERROR_CODE_LENGTH = 4;
    DATE_LENGTH = 8;
    RESPONSE_636_LENGTH = 10;
    RESPONSE_633_LENGTH = 102;
    ERROR_MSG_FILE_NAME = 'ERROR_MSG.txt';

    SEND_TO_DISPATCHER_INFO_DEFAULT_TYPE = 0;
    SEND_TO_DISPATCHER_INFO_TYPE_1 = 1;
    SEND_TO_DISPATCHER_INFO_TYPE_2 = 2;
    SEND_TO_DISPATCHER_INFO_TYPE_3 = 3;

    SOCKET_WAIT_FOR_DATA_TIMEOUT = 30 ;// WaitForData  timeout 计 => 痷タ(WaitForTimeout)莱赣琌把计 frmMain.nG_CmdRefreshRate1 ┪ frmMain.nG_CmdRefreshRate2 (パ畃┪瞒畃丁τ﹚)
    BATCH_OPERATOR = 'уΩ穨';
    DELETE_MAPPING_DATA_TIME = '0200'; // ボ贬 02:00 ,祘Α穦盢 CaSeqNoTransNumMappingData い UPDTIME=琎ぱ data 埃奔
    MAX_COMMAND_COUNT_PER_TIMES = 100;

    DB_LOG_FILE_NAME = 'CaLog.txt';
    MAX_SHOW_MEMO_ERROR_COUNT=1000;

    CA_UI_MAX_COUNT=100;

    TIME_STAMP_FILE_NAME = 'TS.dat';
    THREAD_INFO_FILE = 'ThreadInfo.txt';

    AUTO_INVOKE = 'AUTO_INVOKE';
    CA_INI_FILE_NAME = 'CableNds.ini';
    
    STR_SEP = '''';

    SUPER_USER_ID='sys';
    SUPER_USER_PASSWD='sys';

    FROM_ID_LENGTH = 4;
    TO_ID_LENGTH = 4;
    VERSION_LENGTH = 4;
    
    ACKNOWLEDGE_CMD_ID = 1000;
    NON_ACKNOWLEDGE_CMD_ID = 1001;

    NO_ERROR = '0';
    NO_ERROR_MSG = '';

    COMP_ID_SEP = ','; // new project

    WRITE_TO_MIS_ERROR = 'E013';
    WRITE_TO_MIS_ERROR_MSG = '盢糶misア毖...';

    NOTES_FIELD_FORMAT_ERROR = 'E014';
    NOTES_FIELD_FORMAT_ERROR_MSG = 'NDS001 ぇ Notes 逆Α岿粇...';

    ICC_CARD_NO_LENGTH_ERROR = 'E016';
    ICC_CARD_NO_LENGTH_ERROR_MSG = 'ICC Card Noぇ岿粇...';

    OTHER_ERROR = 'E018';
    OTHER_ERROR_MSG = 'ㄤ岿粇...';

    REAL_ICC_CARD_NO_LENGTH = 10;
    LISTVIEW_SUB_ITEM_COUNT = 8;

    DISCONNECT_CMD = 'DISCONNECT';
    PROCESSOR_SYS_PARAM_FILENAME = 'ProcessorParam.dat';
    DISPATCHER_SYS_PARAM_FILENAME = 'DispatcherParam.dat';
    USER_INFO_FILENAME = 'user.dat';
    PASSWORD_CHECKING_TIMES = 3;
    SEQ_NUM_LENGTH = 8;
    SEND_TRANSACTION_NUM_LENGTH = 9;

    TESTING_CMD_COMP_CODE = '99';    
                    
{
    SYS_PARAM_FILE_NAME = 'PARAM.DAT';

    ACTION_TYPE_1 = 'S';
    ACTION_TYPE_2 = 'A';
    ACTION_TYPE_3 = 'R';
    ACTION_TYPE_4 = 'J';
    ACTION_TYPE_5 = 'Z';

    ACTION_PRIORITY_1 = 'I'; //immediate
    ACTION_PRIORITY_2 = 'G'; //gereral
    ACTION_PRIORITY_3 = 'Z'; //Region Key
    ACTION_PRIORITY_4 = 'H'; //high
    ACTION_PRIORITY_5 = 'L'; //low

    REGULAR_SMS_COMMAND_ID = 1;
    MANAGEMENT_SMS_COMMAND_ID = 0;
    EMMG_ID = 1*16*16;
}

implementation

end.
