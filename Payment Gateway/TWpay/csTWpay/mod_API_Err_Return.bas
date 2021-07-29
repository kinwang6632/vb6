Attribute VB_Name = "mod_API_Err_Return"
Option Explicit

' posapi_txn_Com_951110.zip
' CITI-NPG-POS_API_手冊_V1_06_一般與分期__960115.pdf 中所列
' ===============================================================
    'API錯誤回傳值
    '錯誤代號名稱    回傳值(HEX) 回傳值(DEC) 意義
' ===============================================================
    'HY_R_SOCKET_CONNECT 0x0000000C 12 網路連線失敗
    'HY_R_SOCKET_SEND 0x0000000E 14 資料傳送錯誤
    'HY_R_SOCKET_RECEIVE_TIMEOUT 0x0000000F 15 資料接收逾時
    'HY_R_KEYSTORE_FILE 0x00000010 16 缺少keystore 檔
    'HY_R_HYPOS_LIDM 0x10000001 268435457 訂單編號格式錯誤
    'HY_R_HYPOS_PAN 0x10000002 268435458 信用卡卡號格式錯誤
    'HY_R_HYPOS_EXPDATE 0x10000003 268435459 信用卡有效期限格式錯誤
    'HY_R_HYPOS_PURCHAMT 0x10000004 268435460 金額格式錯誤
    'HY_R_HYPOS_CURRENCY 0x10000005 268435461 幣別格式錯誤
    'HY_R_HYPOS_CAVV 0x10000006 268435462 CAVV 格式錯誤
    'HY_R_HYPOS_RESPONSE 0x10000007 268435463 POS 伺服器回覆的資料有誤
    'HY_R_HYPOS_XID 0x10000008 268435464 交易識別碼格式錯誤
    'HY_R_HYPOS_AUTHRRPID 0x10000009 268435465 授權交易代碼格式錯誤
    'HY_R_HYPOS_ORDER_DESC 0x1000000a 268435466 訂單描述說明格式錯誤
    'HY_R_HYPOS_RECUR 0x1000000b 268435467 recur_num 格式錯誤
    'HY_R_HYPOS_BIRTHDAY 0x1000000c 268435468 birthday 欄位格式錯誤
    'HY_R_HYPOS_PID 0x1000000d 268435469 pid 欄位格式錯誤
    'HY_R_HYPOS_AMT_LIMIT 0x1000000f 268435471 交易金額超過上限
    'HY_R_HYPOS_AUTHCODE 0x10000010 268435472 授權碼欄位格式錯誤
    'HY_R_HYPOS_ORGAMT 0x10000011 268435473 取消金額欄位格式錯誤
    'HY_R_HYPOS_TERMSEQ 0x10000012 268435474 termseq 欄位格式錯誤
    'HY_R_HYPOS_EXPONENT 0x10000013 268435475 幣別欄位格式錯誤
    'HY_R_HYPOS_BATCHID 0x10000014 268435476 批次編號欄位格式錯誤
    'HY_R_HYPOS_BATCHSEQ 0x10000015 268435477 批次序號欄位格式錯誤
    'HY_R_HYPOS_MERID 0x10000016 268435478 merid 欄位格式錯誤
    'HY_R_HYPOS_ECI 0x10000017 268435479 eci 欄位格式錯誤
    'HY_R_HYPOS_RECURDATA 0x10000018 268435480 recur_freq 或recur_end 格式錯誤
' ===============================================================




' ===============================================================
' 底下為舊的元件 HyPos hypos3_1.3.5-win32.exe 中的 DLL 檔 及 PDF 所列
' ===============================================================
    'API錯誤回傳值
    '錯誤代號名稱    回傳值(HEX) 回傳值(DEC) 意義
' ===============================================================
    'HY_R_SOCKET_CONNECT 0x0000000C  12  網路連線失敗
    'HY_R_SOCKET_SEND    0x0000000E  14  資料傳送錯誤
    'HY_R_SOCKET_RECEIVE_TIMEOUT 0x0000000F  15  資料接收逾時
    'HY_R_HYPOS_LIDM 0x10000001  268435457   訂單編號格式錯誤
    'HY_R_HYPOS_PAN      0x10000002  268435458   信用卡卡號格式錯誤
    'HY_R_HYPOS_EXPDATE  0x10000003  268435459   信用卡有效期限格式錯誤
    'HY_R_HYPOS_PURCHAMT 0x10000004  268435460   金額格式錯誤
    'HY_R_HYPOS_CURRENCY 0x10000005  268435461   幣別格式錯誤
    'HY_R_HYPOS_CAVV 0x10000006  268435462   CAVV格式錯誤
    'HY_R_HYPOS_RESPONSE 0x10000007  268435463   POS伺服器回覆的資料有誤
    'HY_R_HYPOS_XID  0x10000008  268435464   交易識別碼格式錯誤
    'HY_R_HYPOS_AUTHRRPID    0x10000009  268435465   授權交易代碼格式錯誤
    'HY_R_HYPOS_ORDER_DESC   0x10000010  268435466   訂單說明格式錯誤
    'HY_R_HYPOS_RECUR    0x10000011  268435467   recur_end、recur_freq或recur_num格式錯誤
    'HY_R_HYPOS_BIRTHDAY 0x10000012  268435468   birthday欄位格式錯誤
    'HY_R_HYPOS_PID  0x10000013  268435469   pid欄位格式錯誤
' ===============================================================
