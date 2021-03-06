
require "dnasp_field.pl";
require "common_field.pl";



##################################################################
#
$	MODULE 		= "sms_protocol";
$	AUTHOR		= "P. Willemin";
$	COPYRIGHT	= "NAGRAVision, (c) 1994-1999";
$	DATE		= "24-APR-1999";
$	VERSION 	= "v1.0";
#
##################################################################

@INCLUDES = (

  '"dnasp_field.h"',
  '<common_field.h>'
);


$MAIN = "SMS_COMMAND";


#------------------------------------------------------------------
#
#  SMS COMMAND FIELDS DEFINITION
#
#------------------------------------------------------------------

$S 	= "STRING_FIELD" ;


%FIELD_TYPE = (

	'int',		'\d+',
        'YN',           '[Y|N]',
        'prodType',     '[S|E|B]',
        'brdMode',	'[N|B]',
        'addressMode',	'[U|G]',
        'lid',		'[a-z]{3,3}',
        'date',         '\d+',
        'time',		'\d+',
        'none',		'.*'
);

%SMS_ERROR_CODES = (
	'0000',		'FATAL_ERROR',
	'0001',		'BAD_ROOT_HEADER_SYNTAX',
	'0002',		'BAD_HEADER_SYNTAX',
	'0003',		'BAD_COMMAND_SYNTAX',
	'0004',		'DATABASE_ERROR',
	'0005',		'MESSAGE_NOT_FOUND',
	'0006',		'PRODUCT_NOT_FOUND',
	'0007',		'CANCELED_CARD',
	'0008',		'UA_NOT_FOUND',
	'0009',		'PPV_IN_THE_PAST',
	'0010',		'STU_ALREADY_EXIST',
	'0011',		'SERVICE_NOT_FOUND',
	'0012',		'TOO_MANY_ITEMS',
	'0013',		'PRODUCT_ALREADY_EXISTS',
	'0014',		'UA_ALREADY_EXISTS',
	'0015',		'BAD_EPG_FORMAT',
	'0016',		'SMS_EVENT_ID_NOT_FOUND',
	'0017',		'PRODUCT_ON_NON_PPV_EVENT',
	'0018',		'EXTERNAL_SYSTEM_ERROR',
	'0019',		'EVENT_ALREADY_IPPV',
	'0020',		'BACKOUT_TYPE_OR_SUBTYPE_NOT_FOUND',
	'0021',		'EVENT_WITH_PPVNB_ALREADY_PROGRAMMED',
	'0022',		'DB_INCONSISTENT',
	'0023',		'MULTIPLE_EVENTS_WITH_SAME_PPVNB_ON_PPV',
	'0024',		'PRODUCT_INCONSISTENT',
	'0025',		'TOO_MANY_ITEMS');

%SMS_ERROR_CODE_EXTENSION = (
	'0000',		'NO_EXTENDED_ERROR_CODE',
	'0001',		'BAD_DEBIT_FORMAT',
	'0002',		'BAD_CREDIT_FORMAT',
       	'0003',		'BAD_CREDIT_MODE',
	'0004',		'BAD_DATE_FORMAT',
	'0005',		'BAD_DATE_SEQUENCE',
	'0006',		'BAD_FREQUENCY_FORMAT',
	'0007',		'BAD_STU_NUMBER_FORMAT',
	'0008',		'BAD_IMS_PRODUCT_ID_FORMAT',
	'0009',		'BAD_SMS_PRODUCT_ID_FORMAT',
	'0010',		'BAD_MESSAGE_NUMBER_FORMAT',
	'0011',		'BAD_PHONE_NUMBER_FORMAT',
	'0012',		'BAD_SMS_EVENT_ID_FORMAT',
	'0013',		'BAD_PRICE_FORMAT',
	'0014',		'BAD_THRESHOLD_CREDIT_FORMAT',
	'0015',		'BAD_UA_FORMAT',
	'0016',		'BAD_ZIP_CODE_FORMAT',
	'0017',		'DIFFERENT_PRODUCTS',
	'0018',		'IDENTICAL_PRODUCTS',
	'0019',		'BAD_BROADCAST_MODE',
	'0020',		'BAD_ADDRESS_TYPE',
	'0021',		'BAD_MOP_PPID',
	'0022',		'BAD_DEST_ID',
	'0023',		'BAD_SOURCE_ID',
	'0024',		'BAD_COMMAND_TYPE',
	'0025',		'BAD_COMMAND_ID',
	'0026',		'BAD_VERSION_FORMAT',
	'0027',		'BAD_NUMBER_FORMAT',
	'0028',		'BAD_FLAG_FORMAT',
	'0029',		'BAD_TIME_FORMAT',
	'0030',		'BAD_RATING_FORMAT',
	'0031',		'BAD_CRC_32',
	'0032',		'BAD_ERROR_CODE',
	'0033',		'BAD_ERROR_CODE_EXT',
	'0034',		'CREDIT_THRESHOLD_TOO_HIGH',
	'0035',		'BAD_PPV_NUMBER_FORMAT',
	'0036',		'BAD_REFERENCE_NUMBER_FORMAT',
	'0037',		'BAD_BLACKOUT_TYPE_FORMAT',
	'0038',		'BAD_NB_OF_SUBTYPES_FORMAT',
	'0039',		'BAD_BLACKOUT_SUBTYPE_FORMAT',
	'0040',		'BAD_SERVICE_UID_FORMAT',
	'0041',		'BAD_SERVICE_NUMBER_FORMAT',
	'0042',		'BAD_TOKEN_NUMBER_FORMAT',
	'0043',		'BAD_EVENT_NUMBER_FORMAT',
	'0044',		'BAD_NUMBER_OF_IPPV_FORMAT',
	'0045',		'BAD_IP_ADDRESS_FORMAT',
	'0046',		'BAD_DEAS_MESSAGE_FORMAT',
	'0047',		'BAD_FIPS_FORMAT',
	'0048',		'EXTERNAL_SYSTEM_NOT_RESPONDING',
	'0049',		'EXTERNAL_SYSTEM_ERROR',
	'0050',		'TOO_MANY_ROWS',
	'0051',		'INVALID_PRODUCT',
	'0052',		'BAD_SERVICE_ID_FORMAT',
	'0053',		'BAD_TRANSPORT_ID_FORMAT',
	'0054',		'BAD_NETWORK_ID_FORMAT',
	'0055',		'BAD_LID_FORMAT',
	'0056',		'BAD_PRIORITY_FORMAT',
	'0057',		'BAD_MODE_FORMAT',
	'0058',		'LENGTH_TOO_LONG',
	'0059',		'BAD_FLAG_VALUE',
	'0060',		'BAD_CC_PORT_FORMAT',
	'0061',		'BAD_TRANSACTION_NUMBER_FORMAT',
	'0062',		'BAD_PURGE_MODE_FORMAT');


%FIELD_DESC = (

	# Field name	Size	Type

	# SMS command header fields

	'TRANSACTION_NUMBER',	"9	$S  int BAD_TRANSACTION_NUMBER_FORMAT",
	'COMMAND_TYPE',		"2	$S  int BAD_COMMAND_TYPE",
	'ACK_COMMAND_TYPE',	"4	$S  int",
        'NACK_STATUS',		"1	$S  int",
	'ERROR_CODE',		"4	$S  int BAD_ERROR_CODE",
	'ERROR_CODE_EXT',	"4	$S  int BAD_ERROR_CODE_EXT",
	'CMDLEN',		"3	$S  int BAD_NUMBER_FORMAT",
	'CMD',			"CMDLEN $S  none",
	'SOURCE_ID',		"4 	$S  int BAD_SOURCE_ID",
	'DEST_ID',		"4	$S  int BAD_DEST_ID",
	'MOP_PPID',		"5	$S  int BAD_MOP_PPID",
	'CREATION_DATE',	"8	$S  date BAD_DATE_FORMAT",
        'BROADCAST_MODE',	"1	$S  brdMode BAD_BROADCAST_MODE",
	'BROADCAST_START_DATE',	"8	$S  date BAD_DATE_FORMAT",
	'BROADCAST_END_DATE',	"8	$S  date BAD_DATE_FORMAT",

	# SMS command address fields

	'ADDRESS_KIND',		"1	$S  addressMode BAD_ADDRESS_TYPE",
	'CMD_ID',		"4	$S  int BAD_COMMAND_ID",
	'UA',			"10	$S  int BAD_UA_FORMAT",

	# SMS command body fields

	'IMS_PRODUCT_ID',	"12	$S  int BAD_IMS_PRODUCT_ID_FORMAT",
	'SMS_PRODUCT_ID',	"12	$S  int BAD_SMS_PRODUCT_ID_FORMAT",
	'BEGIN_DATE',		"8	$S  date BAD_DATE_FORMAT",
	'END_DATE',		"8	$S  date BAD_DATE_FORMAT",
        'STU_NUMBER',		"14	$S  int BAD_STU_NUMBER_FORMAT",
        'THRESHOLD_CREDIT',	"7	$S  int BAD_THRESHOLD_CREDIT_FORMAT",
	'CALL_FREQ',		"2	$S  int BAD_FREQUENCY_FORMAT",
        'CREDIT_LIMIT',		"7	$S  int BAD_CREDIT_FORMAT",
        'IP_ADDRESS',		"15	$S  none BAD_IP_ADDRESS_FORMAT",
	'DATE_FIRST_CALL',	"8	$S  date BAD_DATE_FORMAT",
        'CB_TIME',		"6      $S  time BAD_TIME_FORMAT",
        'CB_DATE',		"8      $S  date BAD_DATE_FORMAT",
        'PPV_NUMBER',		"2	$S  int BAD_PPV_NUMBER_FORMAT",
	'ZIP_CODE',		"5	$S  none BAD_ZIP_CODE_FORMAT",
        'PURCHASE_DATE',	"8	$S  date BAD_DATE_FORMAT",
	'WATCHED_STATUS',	"1	$S  YN BAD_FLAG_FORMAT",
        'EVENT_LEN',		"2	$S  int BAD_NUMBER_FORMAT",
	'EVENT_NAME',		"32     $S  none",	# Fixed length
	'PRICE',		"5	$S  int BAD_PRICE_FORMAT",
        'PHONE_NR1',		"16	$S  int BAD_PHONE_NUMBER_FORMAT",
        'PHONE_NR2',		"16	$S  int BAD_PHONE_NUMBER_FORMAT",
        'PHONE_NR3',		"16	$S  int BAD_PHONE_NUMBER_FORMAT",
        'ABNORMAL_PHONE',	"16	$S  int BAD_PHONE_NUMBER_FORMAT",
        'CC_NUMBER',		"16     $S  int BAD_PHONE_NUMBER_FORMAT",
        'CREDIT_MODE',		"2	$S  int BAD_CREDIT_MODE",
        'CREDIT',		"7	$S  int BAD_CREDIT_FORMAT",
        'DEBIT',		"7	$S  int BAD_DEBIT_FORMAT",
        'RESPONDING',		"1      $S  YN BAD_FLAG_FORMAT",
        'PORT_NUM',		"5      $S  int BAD_CC_PORT_FORMAT",
        'IRD_COMMAND_ID',	"3	$S  int",
	'IRD_OPERATION',	"3	$S  int",
 	'IRD_DATA_LENGTH',	"2	$S  int BAD_NUMBER_FORMAT",
	'IRD_DATA',		"96	$S  none",
        'FORCE_EMM',		"1	$S  YN BAD_FLAG_FORMAT",
        'SUSPEND_ICC',		"1	$S  YN BAD_FLAG_FORMAT",
        'TYPE_OF_PRODUCTS',     "1      $S  prodType BAD_FLAG_FORMAT",
        'NB_OF_PRODUCTS',	"2      $S  int BAD_NUMBER_FORMAT",
        'NETWORK_ID',		"5	$S int BAD_NETWORK_ID_FORMAT",
        'TRANSPORT_ID',		"5	$S int BAD_TRANSPORT_ID_FORMAT",
        'SERVICE_ID',		"5	$S int BAD_SERVICE_ID_FORMAT",
        'LID',			"3	$S lid BAD_LID_FORMAT",
        'PRIORITY',		"1	$S int BAD_PRIORITY_FORMAT",
        'MODE',			"1	$S int BAD_MODE_FORMAT",
        'KIND',			"1	$S int BAD_NUMBER_FORMAT",
        'ID',			"4	$S int BAD_NUMBER_FORMAT",
        'VERSION',		"2	$S int BAD_VERSION_FORMAT",
        'VALIDITY_START_DATE',	"8      $S date BAD_DATE_FORMAT",
        'VALIDITY_END_DATE',	"8      $S date BAD_DATE_FORMAT",
        'VALIDITY_START_TIME',	"8      $S date BAD_TIME_FORMAT",
        'VALIDITY_END_TIME',	"8      $S date BAD_TIME_FORMAT",
        'MSGLEN',		"3	$S int BAD_NUMBER_FORMAT",
        'MESSAGE',		"MSGLEN $S none",
        'PURGE_MODE',		"1	$S int BAD_PURGE_MODE_FORMAT",
        'SMS_EVENT_ID1',	"12	$S int BAD_SMS_EVENT_ID_FORMAT",
        'SMS_EVENT_ID2',	"12	$S int BAD_SMS_EVENT_ID_FORMAT",
        'SMS_EVENT_ID',		"12	$S int BAD_SMS_EVENT_ID_FORMAT",
        'PREVIOUS_START_DATE',  "8	$S date BAD_DATE_FORMAT",
        'PREVIOUS_START_TIME',	"6	$S time BAD_TIME_FORMAT",
	'PREVIOUS_DURATION',	"4	$S int BAD_NUMBER_FORMAT",
        'NEW_START_DATE',  	"8	$S date BAD_DATE_FORMAT",
        'NEW_START_TIME',	"6	$S time BAD_TIME_FORMAT",
	'NEW_DURATION',		"4	$S int BAD_NUMBER_FORMAT",
        'ICC_SUSPENDED',	"1	$S YN BAD_FLAG_FORMAT",
        'REFERENCE_NUMBER',	"4	$S int BAD_REFERENCE_NUMBER_FORMAT",
	'NAME',			"80	$S none",
	'DESCRIPTION',		"250	$S none",
        'EVENT_NUMBER',		"3	$S int BAD_EVENT_NUMBER_FORMAT",
	'SERVICE_UID',		"5	$S int BAD_SERVICE_UID_FORMAT",
        'SERVICE_NUMBER',	"3	$S int BAD_SERVICE_NUMBER_FORMAT",
	'SPECIAL_PPV_EVENT',	"1	$S YN BAD_FLAG_FORMAT",
	'IMPULSE_PURCHASE_ALLOWED',"1	$S YN BAD_FLAG_FORMAT",
	'WATCHED_CRITERION',	"1	$S YN BAD_FLAG_FORMAT",
	'FREE_PREVIEW_TIME',	"2	$S int BAD_TIME_FORMAT",
        'REVERSE_BLACKOUT_FLAG',"1	$S BAD_FLAG_FORMAT",
	'BLACKOUT_TYPE',	"2	$S int BAD_BLACKOUT_TYPE_FORMAT",
	'NB_OF_SUBTYPES',	"3	$S int BAD_NB_OF_SUBTYPES_FORMAT",
	'CLEANUP_DATE',  "8	$S date BAD_DATE_FORMAT",
	'CONDITION_DATE',  "8	$S date BAD_DATE_FORMAT",
	'COLLECT_DATE',  "8	$S date BAD_DATE_FORMAT"
   );


#------------------------------------------------------------------
#
#  ALL DEFINED OBJECT FOR THE SMS COMMAND PROTOCOL
#
#------------------------------------------------------------------

@SMS_COMMAND = (
	        'SMS_CMD_HEADER',
                'SMS_CMD_ADDRESS',
                'SMS_CMD_BODY');


#------------------------------------------------------------------
#
#  SMS COMMAND HEADER DEFINITION
#
#------------------------------------------------------------------


%SMS_CMD_HEADER = (

        'key',  'COMMAND_TYPE',

	'"01"',	'SMS_EMM_HEADER',
        '"02"', 'SMS_CONTROL_HEADER',
        '"03"', 'SMS_PRODUCT_DEF_HEADER',
        '"04"', 'SMS_FEEDBACK_HEADER',
        '"05"', 'SMS_OPERATION_HEADER');


@COMMON_HEADER = ('TRANSACTION_NUMBER','COMMAND_TYPE','SOURCE_ID',
		  'DEST_ID','MOP_PPID','CREATION_DATE');

@SMS_EMM_HEADER = (@COMMON_HEADER,
                   'BROADCAST_MODE','BROADCAST_START_DATE',
                   'BROADCAST_END_DATE','%SMS_CMD_ADDRESS');

@SMS_CONTROL_HEADER = (@SMS_EMM_HEADER);

@SMS_PRODUCT_DEF_HEADER = (@COMMON_HEADER,'%SMS_CMD_BODY');

@SMS_FEEDBACK_HEADER = (@COMMON_HEADER, 'UA', '%SMS_CMD_BODY');

@SMS_OPERATION_HEADER = (@COMMON_HEADER, '%SMS_CMD_BODY');




#------------------------------------------------------------------
#
#  SMS COMMAND ADDRESS DEFINITION
#
#------------------------------------------------------------------


%SMS_CMD_ADDRESS = (

        'key',  'ADDRESS_KIND',

	'"U"',	'SMS_EMM_U',
	'"G"',	'SMS_EMM_G');



@SMS_EMM_U = ('ADDRESS_KIND','UA','%SMS_CMD_BODY');

@SMS_EMM_G = ('ADDRESS_KIND','%SMS_CMD_BODY');



#------------------------------------------------------------------
#
#  SMS COMMAND BODY DEFINITION
#
#------------------------------------------------------------------

%SMS_CMD_BODY = (

	'key',  'CMD_ID',

 	'"0002"', 'SMS_ADD_PRODUCT',
	'"0003"', 'SMS_PRODUCT_RENEWAL',
	'"0004"', 'SMS_PRODUCT_SUSPENSION',
	'"0005"', 'SMS_PRODUCT_REACTIVATION',
	'"0006"', 'SMS_PRODUCT_CANCELLATION',
        '"0007"', 'SMS_ALL_PRODUCT_CANCEL',
        '"0008"', 'SMS_CREDIT_MANAGEMENT',
        '"0009"', 'SMS_UPDATE_CREDIT_THRESHOLD',
        '"0010"', 'SMS_ADD_EVENT_PRODUCT',
	'"0013"', 'SMS_CREATE_CREDIT_FOR_IPURCHASE',
        '"0014"', 'SMS_SUSPEND_IPURCHASE',
        '"0015"', 'SMS_REACTIVATE_IPURCHASE',
        '"0020"', 'SMS_SUSPEND_SUBSCRIBER_ICC',
        '"0021"', 'SMS_REACTIVATE_SUBSCRIBER_ICC',
	'"0048"', 'SMS_SET_ZIP_CODE',
        '"0049"', 'SMS_SET_CALLBACK_PHONE_NR',
        '"0050"', 'SMS_CANCEL_ICC',
        '"0051"', 'SMS_INIT_CARD',
        '"0052"', 'SMS_PAIR_ICC',
        '"0053"', 'SMS_CLEAR_PIN_CODE',
        '"0054"', 'SMS_SET_CALLBACK_IP_ADDRESS',
        '"0060"', 'SMS_IMMEDIATE_CALLBACK',
        '"0061"', 'SMS_ENABLE_AUTO_CALLBACK',
        '"0062"', 'SMS_DISABLE_AUTO_CALLBACK',
        '"0064"', 'SMS_UPDATE_EVENT_RIGHT',
        '"0069"', 'SMS_SEND_GENERIC_IRD',
        '"0071"', 'SMS_GET_PRODUCTS',
        '"0072"', 'SMS_SET_PRODUCTS',
	'"0079"', 'SMS_FORCE_TUNE',
	'"0080"', 'SMS_SEND_MESSAGE',
	'"0092"', 'SMS_PURGE_OLD_PRODUCTS',
	'"0096"', 'SMS_PURGE_PPV_IPPV_RECORDS',
	'"0097"', 'SMS_SET_IPPV_RECORDS_REPORTED',
        '"0100"', 'SMS_CREDIT_LIMIT',
        '"0101"', 'SMS_AUTH_PHONE',
        '"0104"', 'SMS_CREATE_CC_ICC',
        '"0105"', 'SMS_CANCEL_CC_ICC',
        '"0110"', 'SMS_EMM_CLEANUP',

	'"0111"', 'SMS_GET_HISTORY_FROM_CC',

        '"0200"', 'SMS_LOW_CREDIT_ALARM',
        '"0201"', 'SMS_CURRENT_DEBIT_AND_CREDIT',
        '"0202"', 'SMS_PPV_PURCHASE_LIST',
        '"0205"', 'SMS_PHONE_DISCREPANCIES',
        '"0206"', 'SMS_STU_RESPONDING_STATUS',
        '"0207"', 'SMS_ICC_MEMORY_FULL_ALARM',
        '"0208"', 'SMS_EVENT_DEFINITION_ERROR',
        '"0209"', 'SMS_NULL_EVENT_ERROR',
        '"0210"', 'SMS_EPG_DATA_FEED_FORMAT_ERROR',
        '"0211"', 'SMS_START_OF_REPORT',
        '"0212"', 'SMS_END_OF_REPORT',
        '"0213"', 'SMS_EVENT_PRODUCT_SCHEDULE_CHANGE',
        '"0214"', 'SMS_EVENT_REMOVE_ERROR',
        '"0215"', 'SMS_PRODUCT_LIST',

        '"0300"', 'SMS_CREATE_EVENT_PRODUCT',
	'"0301"', 'SMS_REMOVE_PRODUCT',
        '"0302"', 'SMS_MODIFY_EVENT_PRODUCT',
        '"0303"', 'SMS_CREATE_SERVICE_PRODUCT',
        '"0304"', 'SMS_MODIFY_SERVICE_PRODUCT',
        '"0305"', 'SMS_CREATE_SERVICE_PACKAGE_PRODUCT',
        '"0306"', 'SMS_MODIFY_SERVICE_PACKAGE_PRODUCT',
	'"0307"', 'SMS_CREATE_EVENT_PACKAGE_PRODUCT',
        '"0308"', 'SMS_MODIFY_EVENT_PACKAGE_PRODUCT',
	'"0309"', 'SMS_UPDATE_EVENT_PPV_NUMBER',

	'"1000"', 'SMS_ACK',
	'"1001"', 'SMS_NACK',
        '"1002"', 'SMS_NO_COMMAND');


# EMM command
#---------------

#002
@SMS_ADD_PRODUCT = ('CMD_ID','IMS_PRODUCT_ID','BEGIN_DATE','END_DATE');

#003
@SMS_PRODUCT_RENEWAL = ('CMD_ID','IMS_PRODUCT_ID','END_DATE');

#004
@SMS_PRODUCT_SUSPENSION = ('CMD_ID','IMS_PRODUCT_ID');

#005
@SMS_PRODUCT_REACTIVATION = ('CMD_ID','IMS_PRODUCT_ID');

#006
@SMS_PRODUCT_CANCELLATION = ('CMD_ID','IMS_PRODUCT_ID');

#007
@SMS_ALL_PRODUCT_CANCEL = ('CMD_ID');

#008
@SMS_CREDIT_MANAGEMENT = ('CMD_ID','CREDIT_MODE','CREDIT');

#009
@SMS_UPDATE_CREDIT_THRESHOLD = ('CMD_ID','THRESHOLD_CREDIT');

#010
@SMS_ADD_EVENT_PRODUCT = ('CMD_ID','IMS_PRODUCT_ID','EVENT_LEN','EVENT_NAME',
	'PRICE');

#013
@SMS_CREATE_CREDIT_FOR_IPURCHASE = ('CMD_ID', 'CREDIT', 'THRESHOLD_CREDIT');

#014
@SMS_SUSPEND_IPURCHASE = ('CMD_ID');

#015
@SMS_REACTIVATE_IPURCHASE = ('CMD_ID');

#020
@SMS_SUSPEND_SUBSCRIBER_ICC = ('CMD_ID');

#021
@SMS_REACTIVATE_SUBSCRIBER_ICC = ('CMD_ID');

#048
@SMS_SET_ZIP_CODE = ('CMD_ID','ZIP_CODE');

#049
@SMS_SET_CALLBACK_PHONE_NR = ('CMD_ID', 'CC_NUMBER');

#050
@SMS_CANCEL_ICC = ('CMD_ID');

#051
@SMS_INIT_CARD = ('CMD_ID');

#052
@SMS_PAIR_ICC = ('CMD_ID','STU_NUMBER');

#053
@SMS_CLEAR_PIN_CODE = ('CMD_ID');

#054
@SMS_SET_CALLBACK_IP_ADDRESS = ('CMD_ID','IP_ADDRESS','PORT_NUM');




#060
@SMS_IMMEDIATE_CALLBACK = ('CMD_ID');

#061
@SMS_ENABLE_AUTO_CALLBACK = ('CMD_ID','CALL_FREQ','DATE_FIRST_CALL','CB_TIME');

#062
@SMS_DISABLE_AUTO_CALLBACK = ('CMD_ID');

#064
@SMS_UPDATE_EVENT_RIGHT = ('CMD_ID','IMS_PRODUCT_ID','END_DATE');

#069
@SMS_SEND_GENERIC_IRD = ('CMD_ID', 'IRD_COMMAND_ID', 'IRD_OPERATION',
	'IRD_DATA_LENGTH','IRD_DATA');

#071
@SMS_GET_PRODUCTS = ('CMD_ID');

#072
@SMS_SET_PRODUCTS = ('CMD_ID','FORCE_EMM','SUSPEND_ICC','TYPE_OF_PRODUCTS',
	'BEGIN_DATE','END_DATE','NB_OF_PRODUCTS');

#079
@SMS_FORCE_TUNE = ('CMD_ID','NETWORK_ID','TRANSPORT_ID','SERVICE_ID');

#080
@SMS_SEND_MESSAGE = ('CMD_ID','LID','PRIORITY','MODE','KIND','ID',
   	'VERSION','VALIDITY_START_DATE','VALIDITY_END_DATE',
   	'VALIDITY_START_TIME','VALIDITY_END_TIME','MSGLEN',
   	'MESSAGE');

#092
@SMS_PURGE_OLD_PRODUCTS = ('CMD_ID','PURGE_MODE');


#096
@SMS_PURGE_PPV_IPPV_RECORDS = ( 'CMD_ID', 'CLEANUP_DATE', 'CONDITION_DATE' );


#097
@SMS_SET_IPPV_RECORDS_REPORTED = ('CMD_ID','COLLECT_DATE');


# control command
#------------------

#100
@SMS_CREDIT_LIMIT = ('CMD_ID','CREDIT_LIMIT');

#101
@SMS_AUTH_PHONE = ('CMD_ID','PHONE_NR1','PHONE_NR2','PHONE_NR3');

#104
@SMS_CREATE_CC_ICC = ('CMD_ID','STU_NUMBER');

#105
@SMS_CANCEL_CC_ICC = ('CMD_ID');

#110
@SMS_EMM_CLEANUP = ('CMD_ID');

#111
@SMS_GET_HISTORY_FROM_CC = ('CMD_ID');



# Feedback command
#-------------------

#200
@SMS_LOW_CREDIT_ALARM = ('CMD_ID', 'STU_NUMBER', 'CREDIT', 'DEBIT');

#201
@SMS_CURRENT_DEBIT_AND_CREDIT = ('CMD_ID', 'STU_NUMBER', 'CREDIT', 'DEBIT');


#202
@SMS_PPV_PURCHASE_LIST = ('CMD_ID','STU_NUMBER','IMS_PRODUCT_ID',
	'PURCHASE_DATE','WATCHED_STATUS');

#205
@SMS_PHONE_DISCREPANCIES = ('CMD_ID','PHONE_NR1','PHONE_NR2','PHONE_NR3',
       	'ABNORMAL_PHONE');

#206
@SMS_STU_RESPONDING_STATUS = ('CMD_ID','STU_NUMBER','RESPONDING');

#207
@SMS_ICC_MEMORY_FULL_ALARM = ('CMD_ID','STU_NUMBER');

#208
@SMS_EVENT_DEFINITION_ERROR = ('CMD_ID','SMS_EVENT_ID1','SMS_EVENT_ID2');

#209
@SMS_NULL_EVENT_ERROR = ('CMD_ID','SMS_EVENT_ID');

#210
@SMS_EPG_DATA_FEED_FORMAT_ERROR = ('CMD_ID','ERROR_CODE','ERROR_CODE_EXT');

#211
@SMS_START_OF_REPORT = ('CMD_ID', 'CB_DATE', 'CB_TIME');

#212
@SMS_END_OF_REPORT = ('CMD_ID', 'PPV_NUMBER');

#213
@SMS_EVENT_PRODUCT_SCHEDULE_CHANGE = ('CMD_ID','SMS_EVENT_ID',
	'PREVIOUS_START_DATE','PREVIOUS_START_TIME','PREVIOUS_DURATION',
	'NEW_START_DATE','NEW_START_TIME','NEW_DURATION');

#214
@SMS_EVENT_REMOVE_ERROR = ('CMD_ID','SMS_EVENT_ID');

#215
@SMS_PRODUCT_LIST = ('CMD_ID','TRANSACTION_NUMBER','STU_NUMBER','ICC_SUSPENDED',
	'NB_OF_PRODUCTS');



#300
@SMS_CREATE_EVENT_PRODUCT = ('CMD_ID','SMS_PRODUCT_ID','PPV_NUMBER',
	'SMS_EVENT_ID','REFERENCE_NUMBER','NAME','DESCRIPTION',
	'VALIDITY_START_DATE','VALIDITY_START_TIME','VALIDITY_END_DATE',
	'VALIDITY_END_TIME','PRICE','SPECIAL_PPV_EVENT',
	'IMPULSE_PURCHASE_ALLOWED','WATCHED_CRITERION','FREE_PREVIEW_TIME',
	'REVERSE_BLACKOUT_FLAG','BLACKOUT_TYPE','NB_OF_SUBTYPES');

#301
@SMS_REMOVE_PRODUCT = ('CMD_ID','IMS_PRODUCT_ID');

#302
@SMS_MODIFY_EVENT_PRODUCT = ('CMD_ID','IMS_PRODUCT_ID','REFERENCE_NUMBER',
	'VALIDITY_START_DATE','VALIDITY_START_TIME','VALIDITY_END_DATE',
	'VALIDITY_END_TIME','PRICE','SPECIAL_PPV_EVENT',
	'IMPULSE_PURCHASE_ALLOWED','WATCHED_CRITERION','FREE_PREVIEW_TIME',
	'REVERSE_BLACKOUT_FLAG','BLACKOUT_TYPE','NB_OF_SUBTYPES');

#303
@SMS_CREATE_SERVICE_PRODUCT = ('CMD_ID','SMS_PRODUCT_ID','SERVICE_UID',
	'REFERENCE_NUMBER','VALIDITY_START_DATE','VALIDITY_END_DATE',
	'PRICE');

#304
@SMS_MODIFY_SERVICE_PRODUCT = ('CMD_ID','IMS_PRODUCT_ID','SERVICE_UID',
	'REFERENCE_NUMBER','NAME','DESCRIPTION','VALIDITY_START_DATE',
	'VALIDITY_END_DATE','PRICE');

#305
@SMS_CREATE_SERVICE_PACKAGE_PRODUCT = ('CMD_ID','SMS_PRODUCT_ID',
	'REFERENCE_NUMBER','NAME','DESCRIPTION','VALIDITY_START_DATE',
	'VALIDITY_END_DATE','PRICE','SERVICE_NUMBER');

#306
@SMS_MODIFY_SERVICE_PACKAGE_PRODUCT = ('CMD_ID','IMS_PRODUCT_ID',
	'REFERENCE_NUMBER','NAME','DESCRIPTION','VALIDITY_START_DATE',
	'VALIDITY_END_DATE','PRICE','SERVICE_NUMBER');

#307
@SMS_CREATE_EVENT_PACKAGE_PRODUCT = ('CMD_ID','SMS_PRODUCT_ID',
	'REFERENCE_NUMBER','NAME','DESCRIPTION','VALIDITY_START_DATE',
	'VALIDITY_END_DATE','PRICE','EVENT_NUMBER');

#308
@SMS_MODIFY_EVENT_PACKAGE_PRODUCT = ('IMS_PRODUCT_ID','REFERENCE_NUMBER',
	'NAME','DESCRIPTION','VALIDITY_START_DATE','VALIDITY_END_DATE',
	'PRICE','EVENT_NUMBER');


#309
@SMS_UPDATE_EVENT_PPV_NUMBER = ('CMD_ID','SMS_EVENT_ID','PPV_NUMBER');



# SMS command acknowledge
#--------------------------

#1000
@SMS_ACK  = ('CMD_ID','TRANSACTION_NUMBER','IMS_PRODUCT_ID','SMS_PRODUCT_ID');

#1001
@SMS_NACK = ('CMD_ID','TRANSACTION_NUMBER',
    	'NACK_STATUS','ERROR_CODE','ERROR_CODE_EXT','CMDLEN','CMD');

#1002
@SMS_NO_COMMAND = ('ACK_COMMAND_TYPE');




