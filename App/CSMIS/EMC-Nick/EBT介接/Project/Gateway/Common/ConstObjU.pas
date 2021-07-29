unit ConstObjU;

interface

type
    TCompMappingObj = class(TObject)
      EmcCompCode : String;
    end;

const
    SELECT_MODE = 1;
    IUD_MODE = 2;

    ACTIVATE_CAPTION = '�t�ιB�@��';
    DEACTIVATE_CAPTION = '�t�Ω|���Ұ�';

    CONNECT_TO_DB = '��Ʈw�s�u��...';
    CONNECT_DB_OK = '��Ʈw�s�u����';

    DISCONNECT_FROM_DB = '������Ʈw�s�u...';
    DISCONNECT_FROM_DB_OK = '�w�P��Ʈw���u';

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
    ERR_MSG_01 = '�Ȥ�W�٤��i�ť�';
    ERR_CODE_02 = -902;
    ERR_MSG_02 = '�Ȥ��p���q�ܻP�Ȥ��ʹq�ܤ��i���O�ť�';
    ERR_CODE_03 = -903;
    ERR_MSG_03 = '�˾��a�}���i�ť�';
    ERR_CODE_04 = -904;
    ERR_MSG_04 = '�A�ȧO���i�ť�';
    ERR_CODE_05 = -905;
    ERR_MSG_05 = 'CATV Valid���i�ť�';
    ERR_CODE_06 = -906;
    ERR_MSG_06 = 'CATV �����Ĥ�o�S��CATV�Ƚs';
    ERR_CODE_07 = -907;
    ERR_MSG_07 = '�O�_��t�x�������i�ť�';
    ERR_CODE_08 = -908;
    ERR_MSG_08 = 'EBT CATV ID���i�ť�';
    ERR_CODE_09 = -909;
    ERR_MSG_09 = '�X���s�����i�ť�';
    ERR_CODE_10 = -910;
    ERR_MSG_10 = '���ʧǸ����i�ť�';
    ERR_CODE_11 = -911;
    ERR_MSG_11 = '�Ȥ�s�����i�ť�';
    ERR_CODE_12 = -912;
    ERR_MSG_12 = '���ʶ��ؤ��i�ť�';
    ERR_CODE_13 = -913;
    ERR_MSG_13 = '�D��t�x������,CATV �Ƚs���i�ť�';
    ERR_CODE_14 = -914;
    ERR_MSG_14 = '��t�x������,��CATV�Ƚs���i�ť�';

    ERR_CODE_20 = -1000;
    ERR_MSG_20 = '�L�Ī��t�Υx�N�X, �P�_�� EMCCUSTID �e3�X';
    
implementation

end.
