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

    USER_INFO_FILENAME = 'user.dat';

    FORM_INI_FILE_NAME = 'LanguageSetting.ini';

    FORM_INI_SIMPLE_CHINESE = 'FORM_1';

    SYS_INI_FILE_NAME = 'GateWayParam.ini';
    TMP_SYS_INI_FILE_NAME = 'TempGateWayParam.ini';

    DATA_AREA_HEADER = 'COMPINFO';

    SELECT_MODE = 1;
    IUD_MODE = 2;

    SGS_UI_MAX_COUNT=500;
    LISTVIEW_SUB_ITEM_COUNT = 9;

    TESTING_CMD_COMP_CODE = '99';


    EXCHANGE_DATE_QUERY = 'ExchangeDateQuery';
    IC_CARD_QUERY = 'ICCardQuery';
    CA_PRODUCT_QUERY = 'CAProductQuery';
    PROD_PURCHASE_QUERY = 'ProdPurchaseQuery';
    ENTITLEMENT_QUERY ='EntitlementQuery';


    XML_ENCODING_SC = 'GB2312';//簡體
    XML_ENCODING_TC = 'BIG5';//繁體
    XML_STANDALONE = 'yes';


    MAX_FILE_SIZE = 1000;  //大於1MB壓縮
    IIS_HISTORY ='SGS1';
    IIS_IMMEDIATELY ='SGS2';



    SHOW_UI_LINE_COUNTS = 1000;
    SHOW_UI_COUNTS = 1000;
    MAX_STRING_LIST_COUNTS = 10000001;
    //MAX_STRING_LIST_COUNTS = 101;

    ENCRYPTION_SEP = 'H';
    ENCRYPTION_KEY = '1234';
    REGISTRY_ROOT = 'SOFTWARE\Windows\Sys';
    SYS4_Y = 'A';
    SYS4_N = '5';
    SYS5_Y = '9';
    SYS5_N = 'B';

    TAB_STRING = ''#9'';
implementation

end.
