Attribute VB_Name = "mod_API_Err_Return"
Option Explicit

' posapi_txn_Com_951110.zip
' CITI-NPG-POS_API_��U_V1_06_�@��P����__960115.pdf ���ҦC
' ===============================================================
    'API���~�^�ǭ�
    '���~�N���W��    �^�ǭ�(HEX) �^�ǭ�(DEC) �N�q
' ===============================================================
    'HY_R_SOCKET_CONNECT 0x0000000C 12 �����s�u����
    'HY_R_SOCKET_SEND 0x0000000E 14 ��ƶǰe���~
    'HY_R_SOCKET_RECEIVE_TIMEOUT 0x0000000F 15 ��Ʊ����O��
    'HY_R_KEYSTORE_FILE 0x00000010 16 �ʤ�keystore ��
    'HY_R_HYPOS_LIDM 0x10000001 268435457 �q��s���榡���~
    'HY_R_HYPOS_PAN 0x10000002 268435458 �H�Υd�d���榡���~
    'HY_R_HYPOS_EXPDATE 0x10000003 268435459 �H�Υd���Ĵ����榡���~
    'HY_R_HYPOS_PURCHAMT 0x10000004 268435460 ���B�榡���~
    'HY_R_HYPOS_CURRENCY 0x10000005 268435461 ���O�榡���~
    'HY_R_HYPOS_CAVV 0x10000006 268435462 CAVV �榡���~
    'HY_R_HYPOS_RESPONSE 0x10000007 268435463 POS ���A���^�Ъ���Ʀ��~
    'HY_R_HYPOS_XID 0x10000008 268435464 ����ѧO�X�榡���~
    'HY_R_HYPOS_AUTHRRPID 0x10000009 268435465 ���v����N�X�榡���~
    'HY_R_HYPOS_ORDER_DESC 0x1000000a 268435466 �q��y�z�����榡���~
    'HY_R_HYPOS_RECUR 0x1000000b 268435467 recur_num �榡���~
    'HY_R_HYPOS_BIRTHDAY 0x1000000c 268435468 birthday ���榡���~
    'HY_R_HYPOS_PID 0x1000000d 268435469 pid ���榡���~
    'HY_R_HYPOS_AMT_LIMIT 0x1000000f 268435471 ������B�W�L�W��
    'HY_R_HYPOS_AUTHCODE 0x10000010 268435472 ���v�X���榡���~
    'HY_R_HYPOS_ORGAMT 0x10000011 268435473 �������B���榡���~
    'HY_R_HYPOS_TERMSEQ 0x10000012 268435474 termseq ���榡���~
    'HY_R_HYPOS_EXPONENT 0x10000013 268435475 ���O���榡���~
    'HY_R_HYPOS_BATCHID 0x10000014 268435476 �妸�s�����榡���~
    'HY_R_HYPOS_BATCHSEQ 0x10000015 268435477 �妸�Ǹ����榡���~
    'HY_R_HYPOS_MERID 0x10000016 268435478 merid ���榡���~
    'HY_R_HYPOS_ECI 0x10000017 268435479 eci ���榡���~
    'HY_R_HYPOS_RECURDATA 0x10000018 268435480 recur_freq ��recur_end �榡���~
' ===============================================================




' ===============================================================
' ���U���ª����� HyPos hypos3_1.3.5-win32.exe ���� DLL �� �� PDF �ҦC
' ===============================================================
    'API���~�^�ǭ�
    '���~�N���W��    �^�ǭ�(HEX) �^�ǭ�(DEC) �N�q
' ===============================================================
    'HY_R_SOCKET_CONNECT 0x0000000C  12  �����s�u����
    'HY_R_SOCKET_SEND    0x0000000E  14  ��ƶǰe���~
    'HY_R_SOCKET_RECEIVE_TIMEOUT 0x0000000F  15  ��Ʊ����O��
    'HY_R_HYPOS_LIDM 0x10000001  268435457   �q��s���榡���~
    'HY_R_HYPOS_PAN      0x10000002  268435458   �H�Υd�d���榡���~
    'HY_R_HYPOS_EXPDATE  0x10000003  268435459   �H�Υd���Ĵ����榡���~
    'HY_R_HYPOS_PURCHAMT 0x10000004  268435460   ���B�榡���~
    'HY_R_HYPOS_CURRENCY 0x10000005  268435461   ���O�榡���~
    'HY_R_HYPOS_CAVV 0x10000006  268435462   CAVV�榡���~
    'HY_R_HYPOS_RESPONSE 0x10000007  268435463   POS���A���^�Ъ���Ʀ��~
    'HY_R_HYPOS_XID  0x10000008  268435464   ����ѧO�X�榡���~
    'HY_R_HYPOS_AUTHRRPID    0x10000009  268435465   ���v����N�X�榡���~
    'HY_R_HYPOS_ORDER_DESC   0x10000010  268435466   �q�满���榡���~
    'HY_R_HYPOS_RECUR    0x10000011  268435467   recur_end�Brecur_freq��recur_num�榡���~
    'HY_R_HYPOS_BIRTHDAY 0x10000012  268435468   birthday���榡���~
    'HY_R_HYPOS_PID  0x10000013  268435469   pid���榡���~
' ===============================================================
