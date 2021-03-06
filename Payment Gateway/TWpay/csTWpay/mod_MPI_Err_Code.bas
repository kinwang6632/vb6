Attribute VB_Name = "mod_MPI_Err_Code"
    'MPI???~?N?X??
    '1=Root element invalid.
    '5=Format of one or more elements is invalid according to the specification.
    '1126=PARes signature verification failed
    '1157=AAV verification error
    '0=Approved
    '2=Message element not a defined message.
    '3=Required element missing.
    '4=Critical element not recognized.
    '6=Protocol version too old.
    '1003=Invalid Merchant ID/Acquirer BIN
    '1007=Non participating PAN
    '1065=VERes processing error
    '1069=None of the directory servers are available
    '1109=PAN is not enrolled
    '1138=PARes processing error
    '1143=Cardholder could not be authenticated
    '1201=Invalid request
    '1240=The interface is disabled for this merchant
    '1270=Merchant ID does not uniquely identify this account. Please specify Acquirer BIN in order to proceed.
    '1020=Error in processing VEReq at 3-D Secure layer
    '1024=VEReq processing error
    '1061=Error in connection to Visa directory
    '1066=VERes processing error
    '1081=VERes processing error
    '1090=DOM processing error
    '1001=Cannot process payer authentication request
    '1004=A required parameter is missing
    '1005=Invalid PAN
    '1006=Visa directory cache error
    '1008=Cannot process payer authentication request
    '1009=Invalid currency code
    '1010=Invalid transaction ID
    '1014=Invalid merchant transaction stamp
    '1015=Invalid Expiry Date
    '1017=Invalid Password.
    '1018=Invalid Total Amount
    '1019=Invalid Purchase Description
    '1021=Error in processing VEReq at HTTP layer
    '1022=Verify enrolment processing error
    '1025=VEReq processing error
    '1040=Message conversion error
    '1041=Message conversion error
    '1042=Message conversion error
    '1060=Error in connection to Visa directory
    '1062=Error in connection to Visa directory
    '1063=Time out error in connection to Visa directory
    '1067=Connection to Visa directory cannot be closed
    '1080=VERes processing error
    '1084=VERes format error
    '1085=VERes format error
    '1091=DOM processing error
    '1092=SAX processing error
    '
    '1093=DOM processing error
    '1106=Cannot process payer authentication response
    '1111=PAReq processing error
    '1122=PARes processing error
    '1131=Cannot process payer authentication request
    '1136=Cannot process payer authentication request
    '1137=Cannot process payer authentication request
    '1139=HTTP request processing error
    '1094=DOM processing error
    '1100=VERes processing error
    '1101=VERes processing error
    '1104=VERes processing error
    '1105=PARes processing error
    '1107=Cannot process payer authentication response
    '1108=Cannot process payer authentication response
    '1120=PARes processing error
    '1121=PARes processing error
    '1123=PARes processing error
    '1124=PARes processing error
    '1125=PARes processing error
    '1132=Cannot process payer authentication request
    '1133=Cannot process payer authentication request
    '1134=Invalid PARes
    '1140=Cannot process payer authentication request
    '1141=Invalid PARes
    '
    '1142=Invalid PARes
    '1144=PARes verification error
    '1145=Cannot process payer authentication request
    '1151=AAV verification error
    '1153=AAV verification error
    '1161=Cannot process payer authentication request
    '1163=Cannot process payer authentication request
    '1180=DOM processing error
    '1183=SAX processing error
    '1210=3-D Secure protocol error
    '1220=Invalid license key
    '1146=HTTP PUT error
    '1150=AAV verification error
    '1154=AAV verification error
    '1155=Cannot process SPA request
    '1156=Cannot process SPA request
    '1158=A requied parameter is missing in the HTTP request
    '1160=Cannot process payer authentication request
    '1181=DOM processing error
    '1182=DOM processing error
    '1184=DOM processing error
    '1190=PARes timeout
    '1200=A required element was not found
    '1230=Unsupported request
    '1250=Cannot log transaction
    '1260=Cannot be mapped to a valid transaction
    '2001=Configuration file (config.xml) is invalid.
    '2002=File not found
    '2004=Error parsing config.xml
    '2006=Unable to read the keystore
    '2000=Error loading startup info
    '2003=Error reading file
    '2005=Unable to read the key from KeyStore
    '2007=Error reading the Certificate
    '2008=Class not found
    '2009=No such algorithm 2010=Incorrect password
    '3000=Invalid request length
    '3001=Invalid http request
    '3004=Invalid parameter
    '3005=Error in Inflating PARes ver 1.0.x
    '3006=Error in Deflating PAReq ver 1.0.x
    '3009=Session has expired
    '3010=Merchant has no valid license
    '3016=Decryption error
    '3017=Error in connecting to database
    '3018=Error reading keystore from database
    '3002=permission denied
    '3003=Unhandled exception
    '3007=Invalid merchant
    '3008=Error creating Success-Page
    '3011=CRRes processing failed
    '3012=Invalid XML message
    '3013=timed out reading socket
    '3014=Invalid session
    '3015=System currently unavailable, please try again in a few minutes
    '3016=Decryption error
    '3017=Error in connecting to database
    '3018=Error reading keystore from
    '3019=Error saving transaction
    '3020=Database exception
