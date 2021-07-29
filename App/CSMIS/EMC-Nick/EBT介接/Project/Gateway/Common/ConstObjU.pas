unit ConstObjU;

interface

type
    TCompMappingObj = class(TObject)
      EmcCompCode : String;
    end;

const
    SELECT_MODE = 1;
    IUD_MODE = 2;

    ACTIVATE_CAPTION = '系統運作中';
    DEACTIVATE_CAPTION = '系統尚未啟動';

    CONNECT_TO_DB = '資料庫連線中...';
    CONNECT_DB_OK = '資料庫連線完成';

    DISCONNECT_FROM_DB = '結束資料庫連線...';
    DISCONNECT_FROM_DB_OK = '已與資料庫離線';

    PROCESS_FLAG_1='W';
    PROCESS_FLAG_2='P';
    PROCESS_FLAG_3='C';
    PROCESS_FLAG_4='E';
    PROCESS_FLAG_5='R';

    VALIDNESS_COMPCODE = 9999;
    CA_UI_MAX_COUNT = 100;
    LISTVIEW006_SUB_ITEM_COUNT = 39;
    LISTVIEW007_SUB_ITEM_COUNT = 42;
    LISTVIEWSO151_SUB_ITEM_COUNT = 11;
    LISTVIEWSO153_SUB_ITEM_COUNT = 37;
    ORACLE_STR_SEP = '''';
    SUPER_USER_ID='sys';
    SUPER_USER_PASSWD='sys';
    PASSWORD_CHECKING_TIMES = 3;
    USER_INFO_FILENAME = 'user.dat';
    PARAM_INFO_FILENAME = 'param.dat';
    CSIS_INI_FILE_NAME = 'Csis.ini';
    TMP_CSIS_INI_FILE_NAME = 'TmpCsis.ini';    

    ERR_CODE_01 = -901;
    ERR_MSG_01 = '客戶名稱不可空白';
    ERR_CODE_02 = -902;
    ERR_MSG_02 = '客戶聯絡電話與客戶行動電話不可都是空白';
    ERR_CODE_03 = -903;
    ERR_MSG_03 = '裝機地址不可空白';
    ERR_CODE_04 = -904;
    ERR_MSG_04 = '服務別不可空白';
    ERR_CODE_05 = -905;
    ERR_MSG_05 = 'CATV Valid不可空白';
    ERR_CODE_06 = -906;
    ERR_MSG_06 = 'CATV 之有效戶卻沒有CATV客編';
    ERR_CODE_07 = -907;
    ERR_MSG_07 = '是否跨系台移機不可空白';
    ERR_CODE_08 = -908;
    ERR_MSG_08 = 'EBT CATV ID不可空白';
    ERR_CODE_09 = -909;
    ERR_MSG_09 = '合約編號不可空白';
    ERR_CODE_10 = -910;
    ERR_MSG_10 = '異動序號不可空白';
    ERR_CODE_11 = -911;
    ERR_MSG_11 = '客戶編號不可空白';
    ERR_CODE_12 = -912;
    ERR_MSG_12 = '異動項目不可空白';
    ERR_CODE_13 = -913;
    ERR_MSG_13 = '非跨系台移機時,CATV 客編不可空白';
    ERR_CODE_14 = -914;
    ERR_MSG_14 = '跨系台移機時,舊CATV客編不可空白';

    ERR_CODE_20 = -1000;
    ERR_MSG_20 = '無效的系統台代碼, 判斷值 EMCCUSTID 前3碼';
    
implementation

end.
