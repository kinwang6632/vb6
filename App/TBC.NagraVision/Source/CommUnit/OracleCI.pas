// Direct Oracle Access - Oracle Call Interface unit
// Copyright 1997 - 2001 Allround Automations
// support@allroundautomations.nl
// http://www.allroundautomations.nl

{$I Oracle.inc}

// This unit contains the interface to both the OCI7 and new OCI8
// Not all functions are declared, only the ones we needed

unit OracleCI;

interface

{$IFNDEF LINUX}
uses
  WinTypes, Classes, SysUtils, Forms, Controls, Dialogs;
{$ELSE}
uses
  Libc, Classes, SysUtils, QForms, QControls, QDialogs, SyncObjs;
{$ENDIF}

{$IFDEF LINUX}
type
  Bool = LongBool;
  THandle = Pointer;
{$ENDIF}

type // Our own CriticalSection
  TOracleCriticalSection = class(TObject)
  private
    FSection: TRTLCriticalSection;
    Entered: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Enter;
    procedure Leave;
  end;

type // Interface definitions
  sb1 = ShortInt;
  ub1 = Byte;
  eb1 = Byte;
  sb2 = SmallInt;
  ub2 = Word;
  eb2 = Word;
  sb4 = LongInt;
  ub4 = LongInt;
  eb4 = LongInt;
  sword = Integer;
  uword = Integer;
  eword = Integer;
  Tsb2Array = array[0..1000000000] of sb2;
  Psb2Array = ^Tsb2Array;
  Tub4Array = array[0..500000000] of ub4;
  Pub4Array = ^Tub4Array;
  Tub2Array = array[0..1000000000] of ub2;
  Tub1Array = array[0..2000000000] of ub1;
  Pub1Array = ^Tub1Array;
  TPtrArray = array[0..500000000] of Pointer;
  PPtrArray = ^TPtrArray;
  PDouble   = ^Double;
  PPointer  = ^Pointer;

const
  EB4MAXVAL    = 2147483647;
  EB4MINVAL    = 0;
  UB4MAXVAL    = $FFFFFFFF;
  UB4MINVAL    = 0;
  SB4MAXVAL    = 2147483647;
  SB4MINVAL    = -2147483647;
  MINEB4MAXVAL = 2147483647;
  MAXEB4MINVAL = 0;
  MINUB4MAXVAL = $FFFFFFFF;
  MAXUB4MINVAL = 0;
  MINSB4MAXVAL = 2147483647;
  MAXSB4MINVAL = -2147483647;

const // input data types
  SQLT_CHR   = 1;    // (ORANET TYPE) character string
  SQLT_NUM   = 2;    // (ORANET TYPE) oracle numeric
  SQLT_INT   = 3;    // (ORANET TYPE) integer
  SQLT_FLT   = 4;    // (ORANET TYPE) Floating point number
  SQLT_STR   = 5;    // zero terminated string
  SQLT_VNU   = 6;    // NUM with preceding length byte
  SQLT_PDN   = 7;    // (ORANET TYPE) Packed Decimal Numeric
  SQLT_LNG   = 8;    // long
  SQLT_VCS   = 9;    // Variable character string
  SQLT_NON   = 10;   // Null/empty PCC Descriptor entry
  SQLT_RID   = 11;   // rowid
  SQLT_DAT   = 12;   // date in oracle format
  SQLT_VBI   = 15;   // binary in VCS format
  SQLT_BIN   = 23;   // binary data(DTYBIN)
  SQLT_LBI   = 24;   // long binary
  SQLT_UIN   = 68;   // unsigned integer
  SQLT_SLS   = 91;   // Display sign leading separate
  SQLT_LVC   = 94;   // Longer longs (char)
  SQLT_LVB   = 95;   // Longer long binary
  SQLT_AFC   = 96;   // Ansi fixed char
  SQLT_AVC   = 97;   // Ansi Var char
  SQLT_CUR   = 102;  // cursor  type
  SQLT_RDD   = 104;  // rowid descriptor
  SQLT_LAB   = 105;  // label type
  SQLT_OSL   = 106;  // oslabel type
  SQLT_NTY   = 108;  // named object type
  SQLT_REF   = 110;  // ref type
  SQLT_CLOB  = 112;  // character lob
  SQLT_BLOB  = 113;  // binary lob
  SQLT_BFILE = 114;  // binary file lob
  SQLT_CFILE = 115;  // character file lob
  SQLT_RSET  = 116;  // result set type
  SQLT_NCO   = 122;  // named collection type (varray or nested table)
  SQLT_VST   = 155;  // OCIString type
  SQLT_ODT   = 156;  // OCIDate type

const // Modes
  OCI_DEFAULT        = $00; // the default value for parameters and attributes
  OCI_THREADED       = $01; // the application is in threaded environment
  OCI_OBJECT         = $02; // the application is in object environment
  OCI_NON_BLOCKING   = $04; // non blocking mode of operation
  OCI_ENV_NO_MUTEX   = $08; // the environment handle will not be protected by a mutex internally

const // Handle Types (range from 1 - 49)
  OCI_HTYPE_FIRST    = 1;       // start value of handle type
  OCI_HTYPE_ENV      = 1;       // environment handle
  OCI_HTYPE_ERROR    = 2;       // error handle
  OCI_HTYPE_SVCCTX   = 3;       // service handle
  OCI_HTYPE_STMT     = 4;       // statement handle
  OCI_HTYPE_BIND     = 5;       // bind handle
  OCI_HTYPE_DEFINE   = 6;       // define handle
  OCI_HTYPE_DESCRIBE = 7;       // describe handle
  OCI_HTYPE_SERVER   = 8;       // server handle
  OCI_HTYPE_SESSION  = 9;       // authentication handle
  OCI_HTYPE_TRANS    = 10;      // transaction handle
  OCI_HTYPE_COMPLEXOBJECT = 11; // complex object retrieval handle
  OCI_HTYPE_SECURITY = 12;      // security handle
  OCI_HTYPE_SUBSCRIPTION = 13;  // subscription handle
  OCI_HTYPE_DIRPATH_CTX  = 14;  // direct path context
  OCI_HTYPE_DIRPATH_COLUMN_ARRAY = 15;  // direct path column array
  OCI_HTYPE_DIRPATH_STREAM = 16;// direct path stream
  OCI_HTYPE_PROC     = 17;      // process handle
  OCI_HTYPE_LAST     = 17;      // last value of a handle type

const // Descriptor Types (50 - 255)
  OCI_DTYPE_FIRST    = 50;  // start value of descriptor type
  OCI_DTYPE_LOB      = 50;  // lob  locator
  OCI_DTYPE_SNAP     = 51;  // snapshot descriptor
  OCI_DTYPE_RSET     = 52;  // result set descriptor
  OCI_DTYPE_PARAM    = 53;  // a parameter descriptor obtained from ocigparm
  OCI_DTYPE_ROWID    = 54;  // rowid descriptor
  OCI_DTYPE_COMPLEXOBJECTCOMP = 55; // complex object retrieval descriptor
  OCI_DTYPE_FILE     = 56;  // File Lob locator
  OCI_DTYPE_LAST     = 56;  // last value of a descriptor type

const //Object Ptr Types
  OCI_OTYPE_NAME     = 1;   // object name
  OCI_OTYPE_REF      = 2;   // REF to TDO
  OCI_OTYPE_PTR      = 3;   // PTR to TDO

const // Attribute Types
  OCI_ATTR_FNCODE    = 1;
  OCI_ATTR_OBJECT    = 2;
  OCI_ATTR_NONBLOCKING_MODE = 3;
  OCI_ATTR_SQLCODE   = 4;
  OCI_ATTR_ENV       = 5;
  OCI_ATTR_SERVER    = 6;
  OCI_ATTR_SESSION   = 7;
  OCI_ATTR_TRANS     = 8;
  OCI_ATTR_ROW_COUNT = 9;
  OCI_ATTR_SQLFNCODE = 10;
  OCI_ATTR_PREFETCH_ROWS          = 11;
  OCI_ATTR_NESTED_PREFETCH_ROWS   = 12;
  OCI_ATTR_PREFETCH_MEMORY        = 13;
  OCI_ATTR_NESTED_PREFETCH_MEMORY = 14;
  OCI_ATTR_CHAR_COUNT     = 15; // this specifies the bind and define size in characters
  OCI_ATTR_PDSCL          = 16;
  OCI_ATTR_PDFMT          = 17;
  OCI_ATTR_PARAM_COUNT    = 18;
  OCI_ATTR_ROWID          = 19;
  OCI_ATTR_CHARSET        = 20;
  OCI_ATTR_NCHAR          = 21;
  OCI_ATTR_USERNAME       = 22; // username attribute
  OCI_ATTR_PASSWORD       = 23; // password attribute
  OCI_ATTR_STMT_TYPE      = 24; // statement type
  OCI_ATTR_INTERNAL_NAME  = 25;
  OCI_ATTR_EXTERNAL_NAME  = 26;
  OCI_ATTR_XID            = 27;
  OCI_ATTR_TRANS_LOCK     = 28;
  OCI_ATTR_TRANS_NAME     = 29;
  OCI_ATTR_HEAPALLOC      = 30; // memory allocated on the heap
  OCI_ATTR_CHARSET_ID     = 31; // Character Set ID
  OCI_ATTR_CHARSET_FORM   = 32; // Character Set Form
  OCI_ATTR_MAXDATA_SIZE   = 33; // Maximumsize of data on the server
  OCI_ATTR_CACHE_OPT_SIZE = 34; // object cache optimal size
  OCI_ATTR_CACHE_MAX_SIZE = 35; // object cache maximum size percentage
  OCI_ATTR_PINOPTION      = 36; // object cache default pin option
  OCI_ATTR_ALLOC_DURATION = 37; // object cache default allocation duration
  OCI_ATTR_PIN_DURATION   = 38; // object cache default pin duration
  OCI_ATTR_FDO            = 39; // Format Descriptor object attribute
  OCI_ATTR_POSTPROCESSING_CALLBACK = 40;  // Callback to process outbind data
  OCI_ATTR_POSTPROCESSING_CONTEXT  = 41;  // Callback context to process outbind data
  OCI_ATTR_ROWS_RETURNED  = 42;  // Number of rows returned in current iter - for Bind handles
  OCI_ATTR_FOCBK          = 43;  // Failover Callback attribute
  OCI_ATTR_IN_V8_MODE     = 44;  // is the server/service context in V8 mode
  OCI_ATTR_LOBEMPTY       = 45;
  OCI_ATTR_SESSLANG       = 46;  // session language handle

  OCI_ATTR_VISIBILITY         = 47; // visibility
  OCI_ATTR_RELATIVE_MSGID     = 48; // relative message id
  OCI_ATTR_SEQUENCE_DEVIATION = 49; // sequence deviation

  OCI_ATTR_CONSUMER_NAME      = 50; // consumer name
  OCI_ATTR_DEQ_MODE           = 51; // dequeue mode
  OCI_ATTR_NAVIGATION         = 52; // navigation
  OCI_ATTR_WAIT               = 53; // wait
  OCI_ATTR_DEQ_MSGID          = 54; // dequeue message id

  OCI_ATTR_PRIORITY           = 55; // priority
  OCI_ATTR_DELAY              = 56; // delay
  OCI_ATTR_EXPIRATION         = 57; // expiration
  OCI_ATTR_CORRELATION        = 58; // correlation id
  OCI_ATTR_ATTEMPTS           = 59; // # of attempts
  OCI_ATTR_RECIPIENT_LIST     = 60; // recipient list
  OCI_ATTR_EXCEPTION_QUEUE    = 61; // exception queue name
  OCI_ATTR_ENQ_TIME           = 62; // enqueue time (only OCIAttrGet)
  OCI_ATTR_MSG_STATE          = 63; // message state (only OCIAttrGet)
  OCI_ATTR_AGENT_NAME         = 64; // agent name
  OCI_ATTR_AGENT_ADDRESS      = 65; // agent address
  OCI_ATTR_AGENT_PROTOCOL     = 66; // agent protocol

  OCI_ATTR_SENDER_ID          = 68; // sender id
  OCI_ATTR_ORIGINAL_MSGID     = 69; // original message id

  OCI_ATTR_QUEUE_NAME         = 70; // queue name
  OCI_ATTR_NFY_MSGID          = 71; // message id
  OCI_ATTR_MSG_PROP           = 72; // message properties

  OCI_ATTR_NUM_DML_ERRORS     = 73; // num of errs in array DML
  OCI_ATTR_DML_ROW_OFFSET     = 74; // row offset in the array

  OCI_ATTR_DATEFORMAT         = 75; // default date format string
  OCI_ATTR_BUF_ADDR           = 76; // buffer address
  OCI_ATTR_BUF_SIZE           = 77; // buffer size
  OCI_ATTR_DIRPATH_MODE       = 78; // mode of direct path operation
  OCI_ATTR_DIRPATH_NOLOG      = 79; // nologging option
  OCI_ATTR_DIRPATH_PARALLEL   = 80; // parallel (temp seg) option
  OCI_ATTR_NUM_ROWS           = 81; // number of rows in column array
  OCI_ATTR_COL_COUNT          = 82; // columns of column array processed so far.
  OCI_ATTR_STREAM_OFFSET      = 83; // str off of last row processed
  OCI_ATTR_SHARED_HEAPALLOC   = 84; // Shared Heap Allocation Size

  OCI_ATTR_SERVER_GROUP       = 85; // server group name

  OCI_ATTR_MIGSESSION         = 86; // migratable session attribute

  OCI_ATTR_NOCACHE            = 87; // Temporary LOBs

  OCI_ATTR_MEMPOOL_SIZE       = 88; // Pool Size
  OCI_ATTR_MEMPOOL_INSTNAME   = 89; // Instance name
  OCI_ATTR_MEMPOOL_APPNAME    = 90; // Application name
  OCI_ATTR_MEMPOOL_HOMENAME   = 91; // Home Directory name
  OCI_ATTR_MEMPOOL_MODEL      = 92; // Pool Model (proc,thrd,both
  OCI_ATTR_MODES              = 93; // Modes

  OCI_ATTR_SUBSCR_NAME        = 94; // name of subscription
  OCI_ATTR_SUBSCR_CALLBACK    = 95; // associated callback
  OCI_ATTR_SUBSCR_CTX         = 96; // associated callback context
  OCI_ATTR_SUBSCR_PAYLOAD     = 97; // associated payload
  OCI_ATTR_SUBSCR_NAMESPACE   = 98; // associated namespace

  OCI_ATTR_PROXY_CREDENTIALS  = 99; // Proxy user credentials
  OCI_ATTR_INITIAL_CLIENT_ROLES = 100; // Initial client role list

const // Parameter Attribute Types
  OCI_ATTR_UNK                = 101; // unknown attribute
  OCI_ATTR_NUM_COLS           = 102; // number of columns
  OCI_ATTR_LIST_COLUMNS       = 103; // parameter of the column list
  OCI_ATTR_RDBA               = 104; // DBA of the segment header
  OCI_ATTR_CLUSTERED          = 105; // whether the table is clustered
  OCI_ATTR_PARTITIONED        = 106; // whether the table is partitioned
  OCI_ATTR_INDEX_ONLY         = 107; // whether the table is index only
  OCI_ATTR_LIST_ARGUMENTS     = 108; // parameter of the argument list
  OCI_ATTR_LIST_SUBPROGRAMS   = 109; // parameter of the subprogram list
  OCI_ATTR_REF_TDO            = 110; // REF to the type descriptor
  OCI_ATTR_LINK               = 111; // the database link name
  OCI_ATTR_MIN                = 112; // minimum value
  OCI_ATTR_MAX                = 113; // maximum value
  OCI_ATTR_INCR               = 114; // increment value
  OCI_ATTR_CACHE              = 115; // number of sequence numbers cached
  OCI_ATTR_ORDER              = 116; // whether the sequence is ordered
  OCI_ATTR_HW_MARK            = 117; // high-water mark
  OCI_ATTR_TYPE_SCHEMA        = 118; // type's schema name
  OCI_ATTR_TIMESTAMP          = 119; // timestamp of the object
  OCI_ATTR_NUM_ATTRS          = 120; // number of sttributes
  OCI_ATTR_NUM_PARAMS         = 121; // number of parameters
  OCI_ATTR_OBJID              = 122; // object id for a table or view
  OCI_ATTR_PTYPE              = 123; // type of info described by
  OCI_ATTR_PARAM              = 124; // parameter descriptor
  OCI_ATTR_OVERLOAD_ID        = 125; // overload ID for funcs and procs
  OCI_ATTR_TABLESPACE         = 126; // table name space
  OCI_ATTR_PARSE_ERROR_OFFSET = 129; // Parse error offset

const // Credential Types
  OCI_CRED_RDBMS  = 1;   // database username/password
  OCI_CRED_EXT    = 2;   // externally provided credentials

const // Error Return Values
  OCI_SUCCESS           = 0;        // maps to SQL_SUCCESS of SAG CLI
  OCI_SUCCESS_WITH_INFO = 1;        // maps to SQL_SUCCESS_WITH_INFO
  OCI_NO_DATA           = 100;      // maps to SQL_NO_DATA
  OCI_ERROR             = -1;       // maps to SQL_ERROR
  OCI_INVALID_HANDLE    = -2;       // maps to SQL_INVALID_HANDLE
  OCI_NEED_DATA         = 99;       // maps to SQL_NEED_DATA
  OCI_STILL_EXECUTING   = -3123;    // OCI would block error
  OCI_CONTINUE          = -24200;   // Continue with the body of the OCI function

const // Parsing Syntax Types
  OCI_V7_SYNTAX  = 2;  // V7 language
  OCI_V8_SYNTAX  = 3;  // V8 language
  OCI_NTV_SYNTAX = 1;  // Use what so ever is the native lang of server

const // Scrollable Cursor Options
  OCI_FETCH_NEXT     = $02;
  OCI_FETCH_FIRST    = $04;
  OCI_FETCH_LAST     = $08;
  OCI_FETCH_PRIOR    = $10;
  OCI_FETCH_ABSOLUTE = $20;
  OCI_FETCH_RELATIVE = $40;

const // Bind and Define Options
  OCI_SB2_IND_PTR    = $01;
  OCI_DATA_AT_EXEC   = $02;
  OCI_DYNAMIC_FETCH  = $02;
  OCI_PIECEWISE      = $04;

const // piecewise insert/fetch
  OCI_ONE_PIECE      = $00;
  OCI_FIRST_PIECE    = $01;
  OCI_NEXT_PIECE     = $02;
  OCI_LAST_PIECE     = $03;

const // Execution Modes
  OCI_BATCH_MODE        = $01;
  OCI_EXACT_FETCH       = $02;
  OCI_KEEP_FETCH_STATE  = $04;
  OCI_SCROLLABLE_CURSOR = $08;
  OCI_DESCRIBE_ONLY     = $10;
  OCI_COMMIT_ON_SUCCESS = $20;

const // Authentication Modes
  OCI_MIGRATE     = $0001;    // migratable auth context
  OCI_SYSDBA      = $0002;    // for SYSDBA authorization
  OCI_SYSOPER     = $0004;    // for SYSOPER authorization
  OCI_PRELIM_AUTH = $0008;    // for preliminary authorization

const // Piece Information
  OCI_PARAM_IN    = $01;
  OCI_PARAM_OUT   = $02;

const // Transaction Start Flags
  OCI_TRANS_OLD          = $00000000;
  OCI_TRANS_NEW          = $00000001;
  OCI_TRANS_JOIN         = $00000002;
  OCI_TRANS_RESUME       = $00000004;
  OCI_TRANS_STARTMASK    = $000000ff;
  OCI_TRANS_READONLY     = $00000100;
  OCI_TRANS_READWRITE    = $00000200;
  OCI_TRANS_SERIALIZABLE = $00000400;
  OCI_TRANS_ISOLMASK     = $0000ff00;
  OCI_TRANS_LOOSE        = $00010000;
  OCI_TRANS_TIGHT        = $00020000;
  OCI_TRANS_TYPEMASK     = $000f0000;
  OCI_TRANS_NOMIGRATE    = $00100000;
  OCI_TRANS_TWOPHASE     = $01000000; // Transaction End Flags


const // Object Types
  OCI_OTYPE_UNK      = 0;
  OCI_OTYPE_TABLE    = 1;
  OCI_OTYPE_VIEW     = 2;
  OCI_OTYPE_SYN      = 3;
  OCI_OTYPE_PROC     = 4;
  OCI_OTYPE_FUNC     = 5;
  OCI_OTYPE_PKG      = 6;
  OCI_OTYPE_STMT     = 7;

const // Typecodes
  OCI_TYPECODE_REF         = SQLT_REF;      // SQL/OTS OBJECT REFERENCE
  OCI_TYPECODE_DATE        = SQLT_DAT;      // SQL DATE  OTS DATE
  OCI_TYPECODE_SIGNED8     = 27;            // SQL SIGNED INTEGER(8)  OTS SINT8
  OCI_TYPECODE_SIGNED16    = 28;            // SQL SIGNED INTEGER(16)  OTS SINT16
  OCI_TYPECODE_SIGNED32    = 29;            // SQL SIGNED INTEGER(32)  OTS SINT32
  OCI_TYPECODE_REAL        = 21;            // SQL REAL  OTS SQL_REAL
  OCI_TYPECODE_DOUBLE      = 22;            // SQL DOUBLE PRECISION  OTS SQL_DOUBLE
  OCI_TYPECODE_FLOAT       = SQLT_FLT;      // SQL FLOAT(P)  OTS FLOAT(P)
  OCI_TYPECODE_NUMBER      = SQLT_NUM;      // SQL NUMBER(P S)  OTS NUMBER(P S)
  OCI_TYPECODE_DECIMAL     = SQLT_PDN;      // SQL DECIMAL(P S)  OTS DECIMAL(P S)
  OCI_TYPECODE_UNSIGNED8   = SQLT_BIN;      // SQL UNSIGNED INTEGER(8)  OTS UINT8
  OCI_TYPECODE_UNSIGNED16  = 25;            // SQL UNSIGNED INTEGER(16)  OTS UINT16
  OCI_TYPECODE_UNSIGNED32  = 26;            // SQL UNSIGNED INTEGER(32)  OTS UINT32
  OCI_TYPECODE_OCTET       = 245;           // SQL ???  OTS OCTET
  OCI_TYPECODE_SMALLINT    = 246;           // SQL SMALLINT  OTS SMALLINT
  OCI_TYPECODE_INTEGER     = SQLT_INT;      // SQL INTEGER  OTS INTEGER
  OCI_TYPECODE_RAW         = SQLT_LVB;      // SQL RAW(N)  OTS RAW(N)
  OCI_TYPECODE_PTR         = 32;            // SQL POINTER  OTS POINTER
  OCI_TYPECODE_VARCHAR2    = SQLT_VCS;      // SQL VARCHAR2(N)  OTS SQL_VARCHAR2(N)
  OCI_TYPECODE_CHAR        = SQLT_AFC;      // SQL CHAR(N)  OTS SQL_CHAR(N)
  OCI_TYPECODE_VARCHAR     = SQLT_CHR;      // SQL VARCHAR(N)  OTS SQL_VARCHAR(N)
  OCI_TYPECODE_MLSLABEL    = SQLT_LAB;      // OTS MLSLABEL
  OCI_TYPECODE_VARRAY      = 247;           // SQL VARRAY  OTS PAGED VARRAY
  OCI_TYPECODE_TABLE       = 248;           // SQL TABLE  OTS MULTISET
  OCI_TYPECODE_OBJECT      = SQLT_NTY;      // SQL/OTS NAMED OBJECT TYPE
  OCI_TYPECODE_NAMEDCOLLECTION = SQLT_NCO;  // SQL/OTS NAMED COLLECTION TYPE
  OCI_TYPECODE_BLOB        = SQLT_BLOB;     // SQL/OTS BINARY LARGE OBJECT
  OCI_TYPECODE_BFILE       = SQLT_BFILE;    // SQL/OTS BINARY FILE OBJECT
  OCI_TYPECODE_CLOB        = SQLT_CLOB;     // SQL/OTS CHARACTER LARGE OBJECT
  OCI_TYPECODE_CFILE       = SQLT_CFILE;    // SQL/OTS CHARACTER FILE OBJECT

  OCI_TYPECODE_OTMFIRST    = 228;           // first Open Type Manager typecode
  OCI_TYPECODE_OTMLAST     = 320;           // last OTM typecode
  OCI_TYPECODE_SYSFIRST    = 228;           // first OTM system type (internal)
  OCI_TYPECODE_SYSLAST     = 235;           // last OTM system type (internal)

const // Describe Handle Parameter Attributes
  // Attributes common to Columns and Stored Procs
  OCI_ATTR_DATA_SIZE    = 1;
  OCI_ATTR_DATA_TYPE    = 2;
  OCI_ATTR_DISP_SIZE    = 3;
  OCI_ATTR_NAME         = 4;
  OCI_ATTR_PRECISION    = 5;
  OCI_ATTR_SCALE        = 6;
  OCI_ATTR_IS_NULL      = 7;
  OCI_ATTR_TYPE_NAME    = 8;
  OCI_ATTR_SCHEMA_NAME  = 9;
  OCI_ATTR_SUB_NAME     = 10;
  OCI_ATTR_POSITION     = 11;
  // complex object retrieval parameter attributes
  OCI_ATTR_COMPLEXOBJECTCOMP_TYPE       = 50;
  OCI_ATTR_COMPLEXOBJECTCOMP_TYPE_LEVEL	= 51;
  OCI_ATTR_COMPLEXOBJECT_LEVEL          = 52;
  OCI_ATTR_COMPLEXOBJECT_COLL_OUTOFLINE	= 53;
  // Only Columns
  OCI_ATTR_DISP_NAME    = 100;
  // Only Stored Procs
  OCI_ATTR_OVERLOAD     = 210;
  OCI_ATTR_LEVEL        = 211;
  OCI_ATTR_HAS_DEFAULT  = 212;
  OCI_ATTR_IOMODE       = 213;
  OCI_ATTR_RADIX        = 214;
  OCI_ATTR_NUM_ARGS     = 215;
  // only user-defined Type's
  OCI_ATTR_TYPECODE                 = 216;
  OCI_ATTR_COLLECTION_TYPECODE      = 217;
  OCI_ATTR_VERSION                  = 218;
  OCI_ATTR_IS_INCOMPLETE_TYPE       = 219;
  OCI_ATTR_IS_SYSTEM_TYPE           = 220;
  OCI_ATTR_IS_PREDEFINED_TYPE       = 221;
  OCI_ATTR_IS_TRANSIENT_TYPE        = 222;
  OCI_ATTR_IS_SYSTEM_GENERATED_TYPE = 223;
  OCI_ATTR_HAS_NESTED_TABLE         = 224;
  OCI_ATTR_HAS_LOB                  = 225;
  OCI_ATTR_HAS_FILE                 = 226;
  OCI_ATTR_COLLECTION_ELEMENT       = 227;
  OCI_ATTR_NUM_TYPE_ATTRS           = 228;
  OCI_ATTR_LIST_TYPE_ATTRS          = 229;
  OCI_ATTR_NUM_TYPE_METHODS         = 230;
  OCI_ATTR_LIST_TYPE_METHODS        = 231;
  OCI_ATTR_MAP_METHOD               = 232;
  OCI_ATTR_ORDER_METHOD             = 233;
  // only collection element
  OCI_ATTR_NUM_ELEMS       = 234;
  // only type methods
  OCI_ATTR_ENCAPSULATION   = 235;
  OCI_ATTR_IS_SELFISH      = 236;
  OCI_ATTR_IS_VIRTUAL      = 237;
  OCI_ATTR_IS_INLINE       = 238;
  OCI_ATTR_IS_CONSTANT     = 239;
  OCI_ATTR_HAS_RESULT      = 240;
  OCI_ATTR_IS_CONSTRUCTOR  = 241;
  OCI_ATTR_IS_DESTRUCTOR   = 242;
  OCI_ATTR_IS_OPERATOR     = 243;
  OCI_ATTR_IS_MAP          = 244;
  OCI_ATTR_IS_ORDER        = 245;
  OCI_ATTR_IS_RNDS         = 246;
  OCI_ATTR_IS_RNPS         = 247;
  OCI_ATTR_IS_WNDS         = 248;
  OCI_ATTR_IS_WNPS         = 249;

const // ocicpw Modes
  OCI_AUTH = $08;  // Change the password but donot login

const // Other Constants
  OCI_MAX_FNS           = 100;   // max number of OCI Functions
  OCI_SQLSTATE_SIZE     = 5;
  OCI_ERROR_MAXMSG_SIZE = 1024;
  OCI_LOBMAXSIZE        = MINUB4MAXVAL;
  OCI_ROWID_LEN         = 23;
  OCI_FILE_READONLY     = 1;     // readonly mode open for FILE types

const // Fail Over Events
  OCI_FO_END        = $00000001;
  OCI_FO_ABORT      = $00000002;
  OCI_FO_REAUTH     = $00000004;
  OCI_FO_BEGIN      = $00000008;

const // Fail Over Types
  OCI_FO_NONE       = $00000001;
  OCI_FO_SESSION    = $00000002;
  OCI_FO_SELECT     = $00000004;
  OCI_FO_TXNAL      = $00000008;

const // Function Codes
  OCI_FNCODE_INITIALIZE      = 1;
  OCI_FNCODE_HANDLEALLOC     = 2;
  OCI_FNCODE_HANDLEFREE      = 3;
  OCI_FNCODE_DESCRIPTORALLOC = 4;
  OCI_FNCODE_DESCRIPTORFREE  = 5;
  OCI_FNCODE_ENVINIT         = 6;
  OCI_FNCODE_SERVERATTACH    = 7;
  OCI_FNCODE_SERVERDETACH    = 8;
  // unused                    9
  OCI_FNCODE_SESSIONBEGIN    = 10;
  OCI_FNCODE_SESSIONEND      = 11;
  OCI_FNCODE_PASSWORDCHANGE  = 12;
  OCI_FNCODE_STMTPREPARE     = 13;
  OCI_FNCODE_STMTBINDBYPOS   = 14;
  OCI_FNCODE_STMTBINDBYNAME  = 15;
  // unused                    16
  OCI_FNCODE_BINDDYNAMIC     = 17;
  OCI_FNCODE_BINDOBJECT      = 18;
  OCI_FNCODE_BINDARRAYOFSTRUCT = 20;
  OCI_FNCODE_STMTEXECUTE     = 21;
  // unused                    22, 23
  OCI_FNCODE_DEFINE          = 24;
  OCI_FNCODE_DEFINEOBJECT    = 25;
  OCI_FNCODE_DEFINEDYNAMIC   = 26;
  OCI_FNCODE_DEFINEARRAYOFSTRUCT = 27;
  OCI_FNCODE_STMTFETCH       = 28;
  OCI_FNCODE_STMTGETBIND     = 29;
  // unused                    30, 31
  OCI_FNCODE_DESCRIBEANY     = 32;
  OCI_FNCODE_TRANSSTART      = 33;
  OCI_FNCODE_TRANSDETACH     = 34;
  OCI_FNCODE_TRANSCOMMIT     = 35;
  // unused                    36
  OCI_FNCODE_ERRORGET        = 37;
  OCI_FNCODE_LOBOPENFILE     = 38;
  OCI_FNCODE_LOBCLOSEFILE    = 39;
  OCI_FNCODE_LOBCREATEFILE   = 40;
  OCI_FNCODE_LOBDELFILE      = 41;
  OCI_FNCODE_LOBCOPY         = 42;
  OCI_FNCODE_LOBAPPEND       = 43;
  OCI_FNCODE_LOBERASE        = 44;
  OCI_FNCODE_LOBLENGTH       = 45;
  OCI_FNCODE_LOBTRIM         = 46;
  OCI_FNCODE_LOBREAD         = 47;
  OCI_FNCODE_LOBWRITE        = 48;
  // unused                    49
  OCI_FNCODE_SVCCTXBREAK     = 50;
  OCI_FNCODE_SERVERVERSION   = 51;
  // unused                    52, 53
  OCI_FNCODE_ATTRGET         = 54;
  OCI_FNCODE_ATTRSET         = 55;
  OCI_FNCODE_PARAMSET        = 56;
  OCI_FNCODE_PARAMGET        = 57;
  OCI_FNCODE_STMTGETPIECEINFO = 58;
  OCI_FNCODE_LDATOSVCCTX     = 59;
  OCI_FNCODE_RESULTSETTOSTMT = 60;
  OCI_FNCODE_STMTSETPIECEINFO = 61;
  OCI_FNCODE_TRANSFORGET     = 62;
  OCI_FNCODE_TRANSPREPARE    = 63;
  OCI_FNCODE_TRANSROLLBACK   = 64;
  OCI_FNCODE_DEFINEBYPOS     = 65;
  OCI_FNCODE_BINDBYPOS       = 66;
  OCI_FNCODE_BINDBYNAME      = 67;
  OCI_FNCODE_LOBASSIGN       = 68;
  OCI_FNCODE_LOBISEQUAL      = 69;
  OCI_FNCODE_LOBISINIT       = 70;
  OCI_FNCODE_LOBLOCATORSIZE  = 71;
  OCI_FNCODE_LOBCHARSETID    = 72;
  OCI_FNCODE_LOBCHARSETFORM  = 73;
  OCI_FNCODE_LOBFILESETNAME  = 74;
  OCI_FNCODE_LOBFILEGETNAME  = 75;
  OCI_FNCODE_LOGON           = 76;
  OCI_FNCODE_LOGOFF          = 77;

const // LOB Buffering Flush Flags
  OCI_LOB_BUFFER_FREE        = 1;
  OCI_LOB_BUFFER_NOFREE      = 2;

const // OCI Statement Types
  OCI_STMT_SELECT  = 1;  // select statement
  OCI_STMT_UPDATE  = 2;  // update statement
  OCI_STMT_DELETE  = 3;  // delete statement
  OCI_STMT_INSERT  = 4;  // Insert Statement
  OCI_STMT_CREATE  = 5;  // create statement
  OCI_STMT_DROP    = 6;  // drop statement
  OCI_STMT_ALTER   = 7;  // alter statement
  OCI_STMT_BEGIN   = 8;  // begin ... (pl/sql statement)
  OCI_STMT_DECLARE = 9;  // declare .. (pl/sql statement)

const // CHAR/NCHAR/VARCHAR2/NVARCHAR2/CLOB/NCLOB char set "form" information
  SQLCS_IMPLICIT = 1;   // for CHAR, VARCHAR2, CLOB w/o a specified set
  SQLCS_NCHAR    = 2;   // for NCHAR, NCHAR VARYING, NCLOB
  SQLCS_EXPLICIT = 3;   // for CHAR, etc, with "CHARACTER SET ..." syntax
  SQLCS_FLEXIBLE = 4;   // for PL/SQL "flexible" parameters
  SQLCS_LIT_NULL = 5;   // for typecheck of NULL and empty_clob() lits

const // OCI Parameter Types
  OCI_PTYPE_UNK           = 0;    // unknown
  OCI_PTYPE_TABLE	  = 1;    // table
  OCI_PTYPE_VIEW	  = 2;    // view
  OCI_PTYPE_PROC	  = 3;    // procedure
  OCI_PTYPE_FUNC	  = 4;    // function
  OCI_PTYPE_PKG		  = 5;    // package
  OCI_PTYPE_TYPE          = 6;    // user-defined type
  OCI_PTYPE_SYN		  = 7;    // synonym
  OCI_PTYPE_SEQ		  = 8;    // sequence
  OCI_PTYPE_COL		  = 9;    // column
  OCI_PTYPE_ARG		  = 10;   // argument
  OCI_PTYPE_LIST	  = 11;   // list
  OCI_PTYPE_TYPE_ATTR     = 12;   // user-defined type's attribute
  OCI_PTYPE_TYPE_COLL     = 13;   // collection type's element
  OCI_PTYPE_TYPE_METHOD   = 14;   // user-defined type's method
  OCI_PTYPE_TYPE_ARG      = 15;   // user-defined type method's argument
  OCI_PTYPE_TYPE_RESULT   = 16;   // user-defined type method's result

const // OCI List Types
  OCI_LTYPE_UNK           = 0;    // unknown
  OCI_LTYPE_COLUMN        = 1;    // column list
  OCI_LTYPE_ARG_PROC      = 2;    // procedure argument list
  OCI_LTYPE_ARG_FUNC      = 3;    // function argument list
  OCI_LTYPE_SUBPRG        = 4;    // subprogram list
  OCI_LTYPE_TYPE_ATTR     = 5;    // type attribute
  OCI_LTYPE_TYPE_METHOD   = 6;    // type method
  OCI_LTYPE_TYPE_ARG_PROC = 7;    // type method w/o result argument list
  OCI_LTYPE_TYPE_ARG_FUNC = 8;    // type method w/result argument list

const  // OCINumberToInt
  OCI_NUMBER_UNSIGNED     = 0;    // Unsigned type -- ubX
  OCI_NUMBER_SIGNED       = 2;    // Signed type -- sbX

const  // Null indicator information
  OCI_IND_NOTNULL         =  0;   // not NULL
  OCI_IND_NULL            = -1;   // NULL
  OCI_IND_BADNULL         = -2;   // BAD NULL
  OCI_IND_NOTNULLABLE     = -3;   // not NULLable

const  // OCIPinOpt values
  OCI_PIN_DEFAULT         = 1;    // default pin option
  OCI_PIN_ANY             = 3;    // pin any copy of the object
  OCI_PIN_RECENT          = 4;    // pin recent copy of the object
  OCI_PIN_LATEST          = 5;    // pin latest copy of the object

const  // OCILockOpt values
  OCI_LOCK_NONE           = 1;    // null (same as no lock)
  OCI_LOCK_X              = 2;    // exclusive lock

const // OCITypeGetOpt values
  OCI_TYPEGET_HEADER      = 0;
  OCI_TYPEGET_ALL         = 1;

const // OCIDuration values
  OCI_DURATION_BEGIN      = 10;   // beginning sequence of duration
  OCI_DURATION_NULL       = OCI_DURATION_BEGIN - 1; // null duration
  OCI_DURATION_DEFAULT    = OCI_DURATION_BEGIN - 2; // default
  OCI_DURATION_NEXT       = OCI_DURATION_BEGIN - 3; // next special duration
  OCI_DURATION_SESSION    = OCI_DURATION_BEGIN;     // the end of user session
  OCI_DURATION_TRANS      = OCI_DURATION_BEGIN + 1;

const // OCITypeParamMode values
  OCI_TYPEPARAM_IN        = 0;  // in
  OCI_TYPEPARAM_OUT       = 1;  // out
  OCI_TYPEPARAM_INOUT     = 2;  // in-out
  OCI_TYPEPARAM_BYREF     = 3;  // call by reference (implicitly in-out)

const // OCIObjectFree options
  OCI_OBJECTFREE_FORCE    = 1;  // Free object regardless of dirty flag and pincount
  OCI_OBJECTFREE_NONULL   = 2;  // Don't free null structure

const // Direct Path Load constants
  OCI_DIRPATH_LOAD         = 1; // direct path load operation
  OCI_DIRPATH_UNLOAD       = 2; // direct path unload operation
  OCI_DIRPATH_CONVERT      = 3; // direct path convert only operation

  // values for OCI_ATTR_STATE attribute of OCIDirPathCtx
  OCI_DIRPATH_NORMAL       = 1; // can accept rows, last row complete
  OCI_DIRPATH_PARTIAL      = 2; // last row was partial
  OCI_DIRPATH_NOT_PREPARED = 3; // direct path context is not prepared

  // values for cflg argument to OCIDirpathColArrayEntrySet
  OCI_DIRPATH_COL_COMPLETE = 0; // column data is complete
  OCI_DIRPATH_COL_NULL     = 1; // column is null
  OCI_DIRPATH_COL_PARTIAL  = 2; // column data is partial

type // Handle Definitions
  OCIEnv             = pointer;  // OCI environment handle
  OCIError           = pointer;  // OCI error handle
  OCISvcCtx          = pointer;  // OCI service handle
  OCIStmt            = pointer;  // OCI statement handle
  OCIBind            = pointer;  // OCI bind handle
  OCIDefine          = pointer;  // OCI Define handle
  OCIDescribe        = pointer;  // OCI Describe handle
  OCIServer          = pointer;  // OCI Server handle
  OCISession         = pointer;  // OCI Authentication handle
  OCIComplexObject   = pointer;  // OCI COR handle
  OCITrans           = pointer;  // OCI Transaction handle
  OCISecurity        = pointer;  // OCI Security handle
  OCIDirPathCtx      = pointer;  // OCI Direct path handle
  OCIDirPathColArray = pointer;  // OCI Direct path column array
  OCIDirPathStream   = pointer;  // OCI Direct path stream
  OCIDirPathDesc     = pointer;  // OCI Direct path descriptor
  OCIExtProcContext  = pointer;  // OCI External Procedure Context

type // Descriptor Definitions
  OCISnapshot      = pointer;     // OCI snapshot descriptor
  OCIResult        = pointer;     // OCI Result Set Descriptor
  OCILobLocator    = pointer;     // OCI Lob Locator descriptor
  OCIParam         = pointer;     // OCI PARameter descriptor
  OCIComplexObjectComp = pointer; // OCI COR descriptor
  OCIRowid         = pointer;     // OCI ROWID descriptor

type
  OCITypeCode      = ub2;
  OCIType          = pointer;
  OCIDuration      = ub2;
  OCITypeGetOpt    = ub2;
  OCIPinOpt        = ub2;
  OCILockOpt       = ub2;
  OCIInd           = sb2;
  OCIRaw           = Pointer;
  OCIRef           = Pointer;
  OCIColl          = Pointer;
  OCITable         = Pointer;
  OCITypeParamMode = ub2;

  ROCIDate = packed record
    Year: sb2; // gregorian year; range is -4712 <= year <= 9999
    Month,     // month; range is 1 <= month < 12
    Day,       // day; range is 1 <= day <= 31
    Hour,      // hours; range is 0 <= hours <=23
    Min,       // minutes; range is 0 <= minutes <= 59
    Sec: ub1   // seconds; range is 0 <= seconds <= 59
  end;
  OCIDate = ^ROCIDate;

type
  TCDA = packed record
    v2_rc:sb2;                        // V2 return code
    ft:ub2;                           // SQL function type
    rpc:ub4;                          // rows processed count
    peo:ub2;                          // parse error offset
    fc:ub1;                           // OCI function code
    rcs1:ub1;                         // filler area
    rc:ub2;                           // V7 return code
    wrn:ub1;                          // warning flags
    rcs2:ub1;                         // reserved
    rcs3:sword;                       // reserved
    RowIdObject:ub4;                  // Object of RowId (Oracle8)
    RowIdFile:ub2;                    // File of RowId
    rcs6:ub2;
    RowIdBlock:ub4;                   // Block of RowId
    RowIdRecord:ub2;                  // Record of RowId
    ose:sword;                        // OSD dependent error
    chk:ub1;
    rcsp:Pointer;                     // pointer to reserved area
    rcs9:array[44..64] of ub1;        // filler to 64
  end;
  TLDA = TCDA;
  PLDA = ^TCDA;
  THDA = array[0..255] of ub1;


var // "Old" OCI 7 functions
  orlon:function(LDA:PLDA; var HDA:THDA; uid:PChar; uidl:sword;
                 pswd:PChar; pswdl:sword; audit:sword):LongBool; cdecl;
  ologof:function(LDA:PLDA):LongBool; cdecl;
  oerhms:procedure(LDA:PLDA; rcode:sb2; buf:PChar; bufsiz:sword); cdecl;
  oopen:function(var CDA:TCDA; LDA:PLDA; dbn:PChar; dbnl:sword;
                 arsize:sword; uid:PChar; uidl:sword):LongBool; cdecl;
  oparse:function(var CDA:TCDA; sqlstm:PChar; sqllen:sb4;
                  defflg:sword; lngflg:ub4):LongBool; cdecl;
  odefin:function(var CDA:TCDA; pos:sword; var buf; bufl:sword; ftype:sword;
                  scale:sword; var indp; fmt:PChar; fmtl:sword;
                  fmtt:sword; var rlen; rcode:Pointer):LongBool; cdecl;
  oexfet:function(var CDA:TCDA; nrows:ub4; cancel:sword;
                  exact:sword):LongBool; cdecl;
  oexec:function(var CDA:TCDA):LongBool; cdecl;
  oexn:function(var CDA:TCDA; iters:sword; rowoff:sword):LongBool; cdecl;
  ofen:function(var CDA:TCDA; nrows:ub4):LongBool; cdecl;
  oflng:function(var CDA:TCDA; pos:sword; var buf; bufl:sb4; dtype:sword;
                 var retl:ub4; offset:sb4):LongBool; cdecl;
  oclose:function(var CDA:TCDA):LongBool; cdecl;
  odescr:function(var CDA:TCDA; pos:sword; var dbsize:sb4; var dbtype:sb2;
                  cbuf:PChar; var cbufl:sb4; var dsize:sb4; var prec:sb2;
                  var scale:sb2; var nullok:sb2):LongBool; cdecl;
  obreak:function(LDA: PLDA):LongBool; cdecl;
  ocom:function(LDA:PLDA):LongBool; cdecl;
  orol:function(LDA:PLDA):LongBool; cdecl;
  ocon:function(LDA:PLDA):LongBool; cdecl;
  ocof:function(LDA:PLDA):LongBool; cdecl;
  obndra:function(var CDA:TCDA; sqlvar:PChar; sqlvl:sword; var progv;
                  progvl:sword; ftype:sword; scale:sword; indp:Pointer;
                  alen:Pointer; arcode:Pointer; maxsiz:ub4; cursiz:Pointer;
                  fmt:PChar; fmtl:sword; fmtt:sword):LongBool; cdecl;
  odessp:function(LDA:PLDA; objnam:PChar; onlen:Integer;
                  rsv1:Pointer; rsv1ln:Integer; rsv2:Pointer; rsv2ln:Integer;
                  var ovrld, pos, level, argnm, arnlen, dtype, defsup, mode,
                  dtsiz, prec, scale, radix, spare; var arrsiz: ub4):LongBool; cdecl;
  opinit:procedure(mode: ub4); cdecl;
  olog:function(LDA:PLDA; var HDA:THDA; uid:PChar; uidl:sword;
                 pswd:PChar; pswdl:sword; conn:PChar; connl:sword; mode:ub4):LongBool; cdecl;

var // OCI 8.0 functions
  OCIInitialize:
    function(mode: ub4;
             ctxp: pointer;
             malocfp: Pointer;
             ralocfp: Pointer;
             mfreefp: pointer): sword; cdecl;
  OCIHandleAlloc:
    function(parenth:pointer;
             var hndlpp:pointer;
             htype: ub4;
             xtramem_sz: Integer;
             usrmempp: pointer): sword; cdecl;
  OCIHandleFree:
    function(hndlp: pointer;
             hType: ub4): sword; cdecl;
  OCIDescriptorAlloc:
    function(parenth: pointer;
             descpp: pointer; {var}
             dTtype :ub4;
             xtramem_sz: Integer;
             usrmempp: pointer): sword; cdecl;
  OCIDescriptorFree:
    function(descp: pointer;
             dType: ub4): sword; cdecl;
  OCIEnvInit:
    function(var envp: OCIEnv;
             mode: ub4;
             xtramemsz: Integer;
             usrmempp: pointer): sword; cdecl;
  OCIServerAttach:
    function(srvhp: OCIServer;
             errhp: OCIError;
             dblink: PChar;
             dblink_len: sb4;
             mode: ub4): sword; cdecl;
  OCIServerDetach:
    function(srvhp: OCIServer;
             errhp: OCIError;
             mode: ub4): sword; cdecl;
  OCISessionBegin:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             usrhp: OCISession;
             credt: ub4;
             mode: ub4): sword; cdecl;
  OCISessionEnd:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             usrhp: OCISession;
             mode: ub4): sword; cdecl;
  OCILogon:
    function(envhp: OCIEnv;
             errhp: OCIError;
             var svchp: OCISvcCtx;
             username: PChar;
             uname_len: ub4;
             password: PChar;
             password_len: ub4;
             dbname: PChar;
             dbname_len: ub4): sword; cdecl;
  OCILogoff:
    function(svchp: OCISvcCtx;
             errhp: OCIError):sword; cdecl;
  OCIPasswordChange:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             user_name: PChar;
             usernm_len: ub4;
             opasswd: PChar;
             opasswd_len: ub4;
             npasswd: PChar;
             npasswd_len: ub4;
             mode: ub4): sword; cdecl;
  OCIStmtPrepare:
    function(stmtp: OCIStmt;
             errhp: OCIError;
             stmt: PChar;
             stmt_len: ub4;
             language: ub4;
             mode: ub4): sword; cdecl;
  OCIBindByPos:
    function(todo: boolean): sword; cdecl;
//sword   OCIBindByPos  (OCIStmt *stmtp, OCIBind **bindp, OCIError *errhp,
//		       ub4 position, dvoid *valuep, sb4 value_sz,
//		       ub2 dty, dvoid *indp, ub2 *alenp, ub2 *rcodep,
//		       ub4 maxarr_len, ub4 *curelep, ub4 mode);
//
  OCIBindByName:
    function(stmtp: OCIStmt;
             var bindp: OCIBind;
             errhp: OCIError;
             placeholder: PChar;
             placeh_len: sb4;
             valuep: pointer;
             value_sz: sb4;
             dty: ub4;
             indp: pointer;
             alenp: pointer;    {var ub2}
             rcodep: pointer;   {var ub2}
             maxarr_len: ub4;
             curelep: pointer;  {var ub4}
             mode: ub4): sword; cdecl;
  OCIBindObject:
    function(bindp: OCIBind;
             errhp: OCIError;
             tdo: OCIType;
	     pgvpp: Pointer;
             pvszsp: Pointer;
             indpp: Pointer;
	     indszp: Pointer): sword; cdecl;
  OCIBindDynamic:
    function(todo: boolean): sword; cdecl;
//sword   OCIBindDynamic   (OCIBind *bindp, OCIError *errhp, dvoid *ictxp,
//			  OCICallbackInBind icbfp, dvoid *octxp,
//			  OCICallbackOutBind ocbfp);
//
  OCIBindArrayOfStruct:
    function(todo: boolean): sword; cdecl;
//sword   OCIBindArrayOfStruct   (OCIBind *bindp, OCIError *errhp,
//                                ub4 pvskip, ub4 indskip,
//                                ub4 alskip, ub4 rcskip);
//
  OCIStmtGetPieceInfo:
    function(stmtp: OCIStmt;
             errhp: OCIError;
             hndlpp: Pointer;    { receives bind or define handle }
             typep: Pointer;     { var ub4 }
             in_outp: Pointer;   { var ub1 }
             iterp: Pointer;     { var ub4 }
             idxp: Pointer;      { var ub4 }
             piecep: Pointer     { var ub1 }
             ): sword; cdecl;
  OCIStmtSetPieceInfo:
    function(hndlp: Pointer;    { bind or define handle }
             typ: ub4;
             errhp: OCIError;
             bufp: Pointer;     { var buf }
             alenp: Pointer;    { var ub4 }
             piece: ub1;
             indp: Pointer;     { var indp }
             rcodep: Pointer    { var ub2 }
             ): sword; cdecl;
  OCIStmtExecute:
    function(svchp: OCISvcCtx;
             stmtp: OCIStmt;
             errhp: OCIError;
             iters: ub4;
             rowoff: ub4;
             snap_in: OCISnapshot;
             snap_out: OCISnapshot;
             mode: ub4): sword; cdecl;
  OCIDefineByPos:
    function(stmtp: OCIStmt;
             var defnp: OCIDefine;
             errhp: OCIError;
             position: ub4;
             valuep: pointer;
             value_sz: sb4;
             dty: ub2;
             indp: pointer;
             rlenp: pointer; {var ub2}
             rcodep: pointer; {var ub2}
             mode: ub4): sword; cdecl;
  OCIDefineObject:
    function(defnp: OCIDefine;
             errhp: OCIError;
             tdo: Pointer;
             pgvpp: Pointer;
             pvszsp: Pointer;
             indpp: Pointer;
             indszp: Pointer): sword; cdecl;
type
  OCICallbackDefine =
    function(octxp: Pointer;
              defnp: OCIDefine;
              iter: ub4;
              var bufp: Pointer;
              var alenp: Pointer;
              var piece: ub1;
              var indp: Pointer;
              var rcodep: Pointer): sword; cdecl;
var
  OCIDefineDynamic:
    function(defnp: OCIDefine;
             errhp: OCIError;
             octxp: Pointer;
             ocbfp: OCICallbackDefine): sword; cdecl;
  OCIDefineArrayOfStruct:
    function(todo: boolean): sword; cdecl;
//sword   OCIDefineArrayOfStruct  (OCIDefine *defnp, OCIError *errhp, ub4 pvskip,
//                                 ub4 indskip, ub4 rlskip, ub4 rcskip);
//
  OCIStmtFetch:
    function(stmtp: OCIStmt;
             errhp: OCIError;
             nrows: ub4;
             orientation: ub2;
             mode: ub4): sword; cdecl;
  OCIStmtGetBindInfo:
    function(todo: boolean): sword; cdecl;
//sword   OCIStmtGetBindInfo   (OCIStmt *stmtp, OCIError *errhp, ub4 size,
//                              ub4 startloc,
//                              sb4 *found, text *bvnp[], ub1 bvnl[],
//                              text *invp[], ub1 inpl[], ub1 dupl[],
//                              OCIBind *hndl[]);
//
  OCIDescribeAny:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             objptr: Pointer;
             objnm_len: ub4;
             objptr_typ: ub1;
             info_level: ub1;
             objtyp: ub1;
             dschp: OCIDescribe): sword; cdecl;
  OCIParamGet:
    function(hndlp: pointer;
             htype: ub4;
             errhp: OCIError;
             var parmdpp: pointer;
             pos: ub4): sword; cdecl;
  OCIParamSet:
    function(todo: boolean): sword; cdecl;
//sword   OCIParamSet(dvoid *hdlp, ub4 htyp, OCIError *errhp, CONST dvoid *dscp,
//                    ub4 dtyp, ub4 pos);
//
  OCITransStart:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             timeout: uword;
             flags: ub4): sword; cdecl;
  OCITransDetach:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             flags: ub4): sword; cdecl;
  OCITransCommit:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             flags: ub4): sword; cdecl;
  OCITransRollback:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             flags: ub4): sword; cdecl;
  OCITransPrepare:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             flags: ub4): sword; cdecl;
  OCITransForget:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             flags: ub4): sword; cdecl;
  OCIErrorGet:
    function(hndlp: pointer;
             recordno: ub4;
             sqlstate: PChar;
             var errcodep: sb4;
             bufp: PChar;
             bufsiz: ub4;
             eType: ub4): sword; cdecl;
  OCILobAppend:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             dst_locp: OCILobLocator;
             src_locp: OCILobLocator): sword; cdecl;
  OCILobAssign:
    function(envhp: OCIEnv;
             errhp: OCIError;
             src_locp: OCILobLocator;
             var dst_locpp: OCILobLocator): sword; cdecl;
  OCILobCharSetForm:
    function(envhp: OCIEnv;
             errhp: OCIError;
             locp: OCILobLocator;
             var csfrm: ub1): sword; cdecl;
  OCILobCharSetId:
    function(envhp: OCIEnv;
             errhp: OCIError;
             locp: OCILobLocator;
             var csid: ub2): sword; cdecl;
  OCILobCopy:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             dst_locp: OCILobLocator;
             src_locp: OCILobLocator;
             amount: ub4;
             dst_offset: ub4;
             src_offset: ub4): sword; cdecl;
  OCILobDisableBuffering:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator): sword; cdecl;
  OCILobEnableBuffering:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator): sword; cdecl;
  OCILobErase:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator;
             var amount: ub4;
             offset: ub4): sword; cdecl;
  OCILobFileClose:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator): sword; cdecl;
  OCILobFileCloseAll:
    function(todo: boolean): sword; cdecl;
//sword   OCILobFileCloseAll (OCISvcCtx *svchp, OCIError *errhp);
//
  OCILobFileExists:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             filep: OCILobLocator;
             var flag: LongBool): sword; cdecl;
  OCILobFileGetName:
    function(envhp: OCIEnv;
             errhp: OCIError;
             filepp: OCILobLocator;
             dir_alias: PChar;
             var d_length: ub2;
             filename: PChar;
             var f_length: ub2): sword; cdecl;
  OCILobFileIsOpen:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator;
             var flag: LongBool): sword; cdecl;
  OCILobFileOpen:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator;
             mode: ub1): sword; cdecl;
  OCILobFileSetName:
    function(envhp: OCIEnv;
             errhp: OCIError;
             var filepp: OCILobLocator;
             dir_alias: PChar;
             d_length: ub2;
             filename: PChar;
             f_length: ub2): sword; cdecl;
  OCILobFlushBuffer:
    function(svhp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator;
             flag: ub4): sword; cdecl;
  OCILobGetLength:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator;
             var lenp: ub4): sword; cdecl;
  OCILobIsEqual:
    function(todo: boolean): sword; cdecl;
//sword   OCILobIsEqual  (OCIEnv *envhp, CONST OCILobLocator *x,
//                        CONST OCILobLocator *y,
//                        boolean *is_equal);
//
  OCILobLoadFromFile:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             dst_locp: OCILobLocator;
             src_filep: OCILobLocator;
             amount: ub4;
             dst_offset: ub4;
             src_offset: ub4): sword; cdecl;
  OCILobLocatorIsInit:
    function(envhp: OCIEnv;
             errhp: OCIError;
             locp: OCILobLocator;
             var is_initialized: LongBool): sword; cdecl;
  OCILobRead:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator;
             var amtp: ub4;
             offset: ub4;
             bufp: Pointer;
             bufl: ub4;
             ctxp: Pointer;
             cbfp: Pointer;
             csid: ub2;
             csfrm: ub1): sword; cdecl;
  OCILobTrim:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator;
             newlen: ub4): sword; cdecl;
  OCILobWrite:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             locp: OCILobLocator;
             var amtp: ub4;
             offset: ub4;
             bufp: Pointer;
             buflen: ub4;
             piece: ub1;
             ctxp: Pointer;
             cbfp: Pointer;
             csid: ub2;
             csfrm: ub1): sword; cdecl;
  OCIBreak:
    function(hndlp: pointer;
             errhp: OCIError): sword; cdecl;
  OCIServerVersion:
    function(hndlp: pointer;
             errhp: OCIError;
             text: PChar;
             bufsz: ub4;
             hndltype: ub1): sword; cdecl;
  OCIAttrGet:
    function(trgthndlp: pointer;
             trghndltyp: ub4;
             var attributep;
             sizep: pointer; {VAR UB4}
             attrtype: ub4;
             errhp: OCIError): sword; cdecl;
  OCIAttrSet:
    function(trgthndlp: pointer;
             trghndltyp: ub4;
             attributep: pointer;
             size: ub4;
             attrtype: ub4;
             errhp: OCIError): sword; cdecl;
  OCISvcCtxToLda:
    function(svchp: OCISvcCtx;
             errhp: OCIError;
             var LDA: TLDA): sword; cdecl;
  OCILdaToSvcCtx:
    function(var svchp: OCISvcCtx;
             errhp: OCIError;
             var LDA: TLDA): sword; cdecl;
  OCIResultSetToStmt:
    function(todo: boolean): sword; cdecl;
//sword   OCIResultSetToStmt (OCIResult *rsetdp, OCIError *errhp);
//
//
  OCIObjectNew:
    function(envhp: OCIEnv;
             errhp: OCIError;
             svchp: OCISvcCtx;
             typecode: OCITypeCode;
             tdo: OCIType;
             table: Pointer;
             duration: OCIDuration;
             value: LongBool;
             var instance: Pointer): sword; cdecl;
  OCIObjectFree:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer;
             flags: ub2): sword; cdecl;
  OCIObjectPin:
    function(envhp: OCIEnv;
             errhp: OCIError;
             object_ref: OCIRef;
             corhdl: OCIComplexObject;
             pin_option: OCIPinOpt;
             pin_duration: OCIDuration;
             lock_option: OCILockOpt;
             var instance: Pointer): sword; cdecl;
  OCIObjectUnpin:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer): sword; cdecl;
  OCIObjectPinTable:
    function(envhp: OCIEnv;
             errhp: OCIError;
             svchp: OCISvcCtx;
             schema_name: PChar;
             s_n_length: ub4;
             object_name: PChar;
             o_n_length: ub4;
             not_used: Pointer;
             pin_duration: OCIDuration;
             var instance: Pointer): sword; cdecl;
  OCIObjectCopy:
    function(envhp: OCIEnv;
             errhp: OCIError;
             svchp: OCISvcCtx;
             source: Pointer;
             null_source: Pointer;
             target: Pointer;
             null_target: Pointer;
             tdo: OCIType;
             duration: OCIDuration;
             option: ub1): sword; cdecl;
  OCIObjectMarkDelete:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer): sword; cdecl;
  OCIObjectFlush:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer): sword; cdecl;
  OCIObjectGetObjectRef:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer;
             object_ref: OCIRef): sword; cdecl;
  OCIObjectGetTypeRef:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer;
             type_ref: OCIRef): sword; cdecl;
  OCIObjectRefresh:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer): sword; cdecl;
  OCIObjectIsDirty:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer;
             var dirty: LongBool): sword; cdecl;
  OCIObjectIsLocked:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer;
             var lock: LongBool): sword; cdecl;
  OCIObjectGetInd:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer;
             var null_struct: Psb2Array): sword; cdecl;
  OCIObjectExists:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer;
             var exist: LongBool): sword; cdecl;
  OCIObjectGetAttr:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer;
             null_struct: Psb2Array;
             tdo: OCIType;
             names: Pointer;
             lengths: Pointer;
             name_count: ub4;
             indexes: Pointer;
             index_count: ub4;
             var attr_null_status: OCIInd;
             var attr_null_struct: Psb2Array;
             var attr_value: Pointer;
             var attr_tdo: OCIType): sword; cdecl;
  OCIObjectSetAttr:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer;
             null_struct: Psb2Array;
             tdo: OCIType;
             names: Pointer;
             lengths: Pointer;
             name_count: ub4;
             indexes: Pointer;
             index_count: ub4;
             attr_null_status: OCIInd;
             attr_null_struct: Psb2Array;
             attr_value: Pointer): sword; cdecl;
  OCIObjectLock:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer): sword; cdecl;
  OCIObjectMarkUpdate:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer): sword; cdecl;
  OCIObjectUnmark:
    function(envhp: OCIEnv;
             errhp: OCIError;
             instance: Pointer): sword; cdecl;
  OCITypeByName:
    function(envhp: OCIEnv;
             errhp: OCIError;
             svchp: OCISvcCtx;
             schema_name: PChar;
             s_length: ub4;
             type_name: PChar;
             t_length: ub4;
             version_name: PChar;
             v_length: ub4;
             pin_duration: OCIDuration;
             get_option: OCITypeGetOpt;
             var tdo: OCIType): sword; cdecl;

   OCIStringPtr:
     function(envhp: OCIEnv;
              vs: Pointer): PChar; cdecl;
   OCIStringAssignText:
     function(envhp: OCIEnv;
              errhp: OCIError;
              rhs: PChar;
              rhs_len: ub4;
              var lhs: Pointer): sword; cdecl;
   OCINumberToInt:
     function(errhp: OCIError;
              number: Pointer;
              rsl_length: uword;
              rsl_flag: uword;
              var rsl): sword; cdecl;
   OCINumberFromInt:
     function(errhp: OCIError;
              inum: Pointer;
              inum_length: uword;
              inum_s_flag: uword;
              number: Pointer): sword; cdecl;
   OCINumberToReal:
     function(errhp: OCIError;
              number: Pointer;
              rsl_length: uword;
              var rsl): sword; cdecl;
   OCINumberToText:
     function(errhp: OCIError;
              number: Pointer;
              fmt: PChar;
              fmt_length: ub4;
              nls_params: PChar;
              nls_p_length: ub4;
              var buf_size: ub4;
              buf: PChar): sword; cdecl;
   OCINumberFromText:
     function(errhp: OCIError;
              str: PChar;
              str_length: ub4;
              fmt: PChar;
              fmt_length: ub4;
              nls_params: PChar;
              nls_p_length: ub4;
              number: Pointer): sword; cdecl;
   OCINumberFromReal:
     function(errhp: OCIError;
              rnum: Pointer;
              rnum_length: uword;
              number: Pointer): sword; cdecl;
   OCIRawPtr:
     function(envhp: OCIEnv;
              raw: OCIRaw): Pub1Array; cdecl;
   OCIRawSize:
     function(envhp: OCIEnv;
              raw: OCIRaw): ub4; cdecl;
   OCIRawAssignBytes:
     function(envhp: OCIEnv;
              errhp: OCIError;
              rhs: Pub1Array;
              rhs_len: ub4;
              var raw: OCIRaw): sword; cdecl;
   OCIRefAssign:
     function(envhp: OCIEnv;
              errhp: OCIError;
              source: OCIRef;
              var target: OCIRef): sword; cdecl;
   OCIRefClear:
     function(envhp: OCIEnv;
              ref: OCIRef): LongBool; cdecl;
   OCIRefIsNull:
     function(envhp: OCIEnv;
              ref: OCIRef): LongBool; cdecl;
   OCIRefHexSize:
     function(envhp: OCIEnv;
              ref: OCIRef): ub4; cdecl;
   OCIRefToHex:
     function(envhp: OCIEnv;
              errhp: OCIError;
              ref: OCIRef;
              hex: PChar;
              var hex_length: ub4): sword; cdecl;
   OCIRefFromHex:
     function(envhp: OCIEnv;
              errhp: OCIError;
              svchp: OCISvcCtx;
              hex: PChar;
              hex_length: ub4;
              var ref: OCIRef): sword; cdecl;
(*
   OCIDateGetDate:
     procedure(date: Pointer;
               var year: sb2;
               var month: ub1;
               var day: ub1); cdecl;
   OCIDateGetTime:
     procedure(date: Pointer;
               var Hour: ub1;
               var Min: ub1;
               var Sec: ub1); cdecl;
*)
   OCICollSize:
     function(envhp: OCIEnv;
              errhp: OCIError;
              coll: OCIColl;
              var size: sb4): sword; cdecl;
   OCICollTrim:
     function(envhp: OCIEnv;
              errhp: OCIError;
              trim_num: sb4;
              coll: OCIColl): sword; cdecl;
   OCICollMax:
     function(envhp: OCIEnv;
              coll: OCIColl): sb4; cdecl;
   OCICollGetElem:
     function(envhp: OCIEnv;
              errhp: OCIError;
              coll: OCIColl;
              index: sb4;
              var exists: LongBool;
              var elem: Pointer;
              var elemind: Psb2Array): sword; cdecl;
   OCICollAssignElem:
     function(envhp: OCIEnv;
              errhp: OCIError;
              index: sb4;
              elem: Pointer;
              elemind: Psb2Array;
              coll: OCIColl): sword; cdecl;
   OCICollAppend:
     function(envhp: OCIEnv;
              errhp: OCIError;
              elem: Pointer;
              elemind: Psb2Array;
              coll: OCIColl): sword; cdecl;
   OCITableDelete:
     function(envhp: OCIEnv;
              errhp: OCIError;
              index: sb4;
              tbl: OCITable): sword; cdecl;
type
   OCICacheFlushGet =
     function(context: Pointer;
              var last: ub1): OCIRef; cdecl;
var
   OCICacheFlush:
     function(envhp: OCIEnv;
              errhp: OCIError;
              svchp: OCISvcCtx;
              context: Pointer;
              get: OCICacheFlushGet;
              ref: Pointer): sword; cdecl;
   OCIExtProcGetEnv:
     function(with_context: OCIExtProcContext;
              var envh: OCIEnv;
              var svch: OCISvcCtx;
              var errh: OCIError): sword; cdecl;
   OCIExtProcRaiseExcpWithMsg:
     function(with_context: OCIExtProcContext;
              errnum: Integer;
              errmsg: PChar;
              msglen: Integer): sword; cdecl;
   OCIEnvCreate:
     function(var envhp: OCIEnv;
              mode: ub4;
              ctxp: Pointer;
              malocfp: Pointer;
              ralocfp: Pointer;
              mfreefp: Pointer;
              xtramemsz: Integer;
              usrmempp: Pointer): sword; cdecl;
   OCIDirPathAbort:
     function(dpctx: OCIDirPathCtx;
              errhp: OCIError): sword; cdecl;
//   OCIDirPathDataSave(OCIDirPathCtx *dpctx, OCIError *errhp);
   OCIDirPathFinish:
     function(dpctx: OCIDirPathCtx;
              errhp: OCIError): sword; cdecl;
//   OCIDirPathFlushRow(OCIDirPathCtx *dpctx, OCIError *errhp);
   OCIDirPathPrepare:
     function(dpctx: OCIDirPathCtx;
               svchp: OCISvcCtx;
               errhp: OCIError): sword; cdecl;
   OCIDirPathLoadStream:
     function(dpctx: OCIDirPathCtx;
              dpstr: OCIDirPathStream;
              errhp: OCIError): sword; cdecl;
//   OCIDirPathColArrayEntryGet(OCIDirPathColArray *dpca, OCIError *errhp,
//                              ub4 rownum, ub2 colIdx, ub1 **cvalpp, ub4 *clenp,
//                              ub1 *cflgp);
   OCIDirPathColArrayEntrySet:
     function(dpca: OCIDirPathColArray;
              errhp: OCIError;
              rownum: ub4;
              colIdx: ub2;
              cvalp: Pointer;
              clen: ub4;
              cflg: ub1): sword; cdecl;
//   OCIDirPathColArrayRowGet(OCIDirPathColArray *dpca, OCIError *errhp,
//                            ub4 rownum, ub1 ***cvalppp, ub4 **clenpp,
//                            ub1 **cflgpp);
//   OCIDirPathColArrayReset:
//     function(dpca: OCIDirPathColArray;
//              errhp: OCIError): sword; cdecl;
   OCIDirPathColArrayToStream:
     function(dpca: OCIDirPathColArray;
              dpctx: OCIDirPathCtx;
              dpstr: OCIDirPathStream;
              errhp: OCIError;
              rowcnt: ub4;
              rowoff: ub4): sword; cdecl;
   OCIDirPathStreamReset:
     function(dpstr: OCIDirPathStream;
              errhp: OCIError): sword; cdecl;
//   OCIDirPathStreamToStream(OCIDirPathStream *istr,  OCIDirPathStream *ostr,
//                            OCIDirPathCtx    *dpctx, OCIError         *errhp,
//                            ub4               isoff, ub4               osoff);

// Oracle MTS

const // various error codes reported by the OraMTS functions
  // Dispenser object related errors
  ORAMTSERR_NOERROR     = 0;     // success code
  ORAMTSERR_NOMTXDISPEN = 1001;  // no MTXDM.DLL available
  ORAMTSERR_DSPCREAFAIL = 1002;  // failure to create dispen
  ORAMTSERR_DSPMAXSESSN = 1003;  // exceeded max sessions
  ORAMTSERR_DSPINVLSVCC = 1004;  // invalid OCI Svc ctx
  ORAMTSERR_DSPNODBIDEN = 1005;  // can't create new dbiden
  // Database identifier related errors
  ORAMTSERR_NOSERVEROBJ = 2001;  // unable to alloc a server*/
  // server object related errors
  ORAMTSERR_INVALIDSRVR = 3001;  // invalid server object   */
  ORAMTSERR_FAILEDATTCH = 3002;  // failed attach to Oracle
  ORAMTSERR_FAILEDDETCH = 3003;  // failed detach from db
  ORAMTSERR_FAILEDTRANS = 3004;  // failed to start trans.
  ORAMTSERR_SETATTRIBUT = 3005;  // OCI set attrib failed
  ORAMTSERR_CONNXBROKEN = 3006;  // conn to Oracle broken
  ORAMTSERR_NOTATTACHED = 3007;  // not attached to Oracle
  ORAMTSERR_ALDYATTACHD = 3008;  // alrdy attached to Oracle
  // Session object related errors
  ORAMTSERR_INVALIDSESS = 4001;  // invalid session object
  ORAMTSERR_FAILEDLOGON = 4002;  // failed logon to Oracle
  ORAMTSERR_FAILEDLOGOF = 4003;  // failed logoff from db
  ORAMTSERR_TRANSEXISTS = 4004;  // no transaction beneath
  ORAMTSERR_LOGONEXISTS = 4005;  // already logged on to db
  ORAMTSERR_NOTLOGGEDON = 4006;  // not logged on to Oracle
  // RPC errors
  ORAMTSERR_RPCINVLCTXT = 5001;  // RPC context is invalid
  ORAMTSERR_RPCCOMMUERR = 5002;  // generic communic. error
  ORAMTSERR_RPCALRDYCON = 5003;  // endpoint already connect
  ORAMTSERR_RPCNOTCONNE = 5004;  // endpoint not connected
  ORAMTSERR_RPCPROTVIOL = 5005;  // protocol violation
  ORAMTSERR_RPCACCPTIMO = 5006;  // timeout accepting conn.
  ORAMTSERR_RPCILLEGOPC = 5007;  // invalid RPC opcode
  ORAMTSERR_RPCBADINCNO = 5008;  // mismatched incarnation#
  ORAMTSERR_RPCCONNTIMO = 5009;  // client connect timeout
  ORAMTSERR_RPCSENDTIMO = 5010;  // synch. send timeout
  ORAMTSERR_RPCRECVTIMO = 5011;  // synch. receive timedout
  ORAMTSERR_RPCCONRESET = 5012;  // connection reset by peer
  // Miscellaneous errors
  ORAMTSERR_INVALIDARGU = 6001;  // invalid args to function
  ORAMTSERR_INVALIDOBJE = 6002;  // an object was invalid
  ORAMTSERR_ILLEGALOPER = 6003;  // illegal operation
  ORAMTSERR_ALLOCMEMORY = 6004;  // memory allocation error
  ORAMTSERR_ERRORSYNCHR = 6005;  // synchr. object error
  ORAMTSERR_NOORAPROXY  = 6006;  // no Oracle Proxy server
  ORAMTSERR_ALRDYENLIST = 6007;  // session already enlisted
  ORAMTSERR_NOTENLISTED = 6008;  // session is not enlisted
  ORAMTSERR_TYPMANENLIS = 6009;  // illeg on manuenlst sess
  ORAMTSERR_TYPAUTENLIS = 6010;  // illeg on autoenlst sess
  ORAMTSERR_TRANSDETACH = 6011;  // error detaching trans.
  ORAMTSERR_OCIHNDLALLC = 6012;  // OCI handle alloc error
  ORAMTSERR_OCIHNDLRELS = 6013;  // OCI handle dealloc error
  ORAMTSERR_TRANSEXPORT = 6014;  // error exporting trans.

const // Connection flags on call OraMTSSvcGet()
  ORAMTS_CFLG_ALLDEFAULT = $00;  // default flags
  ORAMTS_CFLG_NOIMPLICIT = $01;  // don't do implicit enlistment
  ORAMTS_CFLG_UNIQUESRVR = $02;  // need a separate Net8 connect
  ORAMTS_CFLG_SYSDBALOGN = $04;  // logon as a SYSDBA
  ORAMTS_CFLG_SYSOPRLOGN = $10;  // logon as a SYSOPER
  ORAMTS_CFLG_PRELIMAUTH = $20;  // preliminary internal login

const // Enlistment flags on call kpntsvcenlist()
  ORAMTS_ENFLG_DEFAULT = 0;      // default flags
  ORAMTS_ENFLG_RESUMTX = 1;      // resume a detached transact.
  ORAMTS_ENFLG_DETCHTX = 2;      // detached from the transact.

var
  OraMTSSvcGet:
    function(lpUName: PChar;
             lpPsswd: PChar;
             lpDbnam: PChar;
             var pOCISvc: OCISvcCtx;
             var pOCIEnv: OCIEnv;
             ConFlg: ub4): sword; cdecl;
  OraMTSSvcRel:
    function(OCISvc: OCISvcCtx): sword; cdecl;
  OraMTSSvcEnlist:
    function(OCISvc: OCISvcCtx;
             OCIErr: OCIError;
             lpTrans: pointer;
             dwFlags: LongInt): sword; cdecl;
  OraMTSSvcEnlistEx:
    function(OCISvc: OCISvcCtx;
             OCIErr: OCIError;
             lpTrans: pointer;
             dwFlags: LongInt;
             lpDBName: PChar): sword; cdecl;
  OraMTSTransTest:
    function: Bool; cdecl;

var
  OCIDLL: string = '';                 // Name of OCI DLL
  MTSDLL: string = '';                 // Path to DLL
  DefaultOCIDLL: Boolean = True;       // Was default OCI DLL used?
  ExcludedOCIDLLs: string = '';        // Comma separated list of excluded OCI DLL's
{$IFDEF LINUX}
  HDLL: THandle = nil;
  HMTS: THandle = nil;
{$ELSE}
  HDLL: THandle = hInstance_Error - 1; // Handle of loaded OCI DLL
  HMTS: THandle = hInstance_Error - 1; // Handle of loaded MTS DLL
{$ENDIF}
  ForceOCI7: Boolean = False;          // Force OCI7 mode
  OCI70: Boolean = False;
  OCI72: Boolean = False;
  OCI73: Boolean = False;
  OCI80: Boolean = False;
  OCI81: Boolean = False;
  OCI80Detected:   Boolean = False;
  OCI81Detected:   Boolean = False;
  ExtProcDetected: Boolean = False;
  MTSDetected:     Boolean = False;
  OracleHomeKey:   string = '';        // Oracle registry key
  OracleHomeDir:   string = '';        // ORACLE_HOME directory
  OracleHomeName:  string = '';        // Forced Oracle home (name)
  DefaultHomeName: Boolean = True;     // Was default OracleHome used?
  InitOCILog:      string = '';        // InitOCI logging
  OCISection: TOracleCriticalSection = nil;  // Critical OCI section

const // Result of DLLInit
  dllOK         = 0;
  dllNoFile     = 1;    // No dll file found
  dllMismatch   = 2;    // dll was no oci dll
  dllNoRegistry = 3;    // No Oracle registry entry found

// Exported functions

{$IFNDEF LINUX}
function  ReadRegString(Root: HKEY; const Key, Name: String): string;
{$ENDIF}
procedure InitOCI;
function  DLLLoaded: Boolean;
function  DLLInit: Integer;
procedure DLLExit;
procedure ResetOCIPath;
function  OCIVersion: string;
procedure BuildOracleAliasList;
procedure BuildOracleHomeList;
function  OracleHome: string;
function  GetOCIDLLList(DoLog: Boolean): TStringList;

procedure InitMTS;
function  MTSLoaded: Boolean;

function TNSNames: string;
function OracleAliasList: TStringList;
function OracleHomeList: TStringList;

implementation

var
  OriginalPath: string = '';                 // DefaultPath;
  OracleTNSNames: string = '';               // Full Path to the tnsnames.ora file
  AliasList: TStringList = nil;              // List of Oracle Aliases
  HomeList:  TStringList = nil;              // List of Oracle Homes

// TOracleCriticalSection

constructor TOracleCriticalSection.Create;
begin
  inherited Create;
  Entered := 0;
  InitializeCriticalSection(FSection);
end;

destructor TOracleCriticalSection.Destroy;
begin
  DeleteCriticalSection(FSection);
  inherited Destroy;
end;

procedure TOracleCriticalSection.Enter;
begin
  EnterCriticalSection(FSection);
  Inc(Entered);
end;

procedure TOracleCriticalSection.Leave;
begin
  if Entered > 0 then
  begin
    Dec(Entered);
    LeaveCriticalSection(FSection);
  end;
end;

// Get a commandline parameter by name
function GetParamString(Name: string): string;
var i: Integer;
begin
  for i := 1 to ParamCount do
  begin
    if Pos(AnsiUpperCase(Name) + '=', AnsiUpperCase(ParamStr(i))) > 0 then
    begin
      Result := Copy(ParamStr(i), Length(Name) + 2, 255);
      Exit;
    end;
  end;
  Result := '';
end;

{$IFNDEF LINUX}
// Read a string from the registry
function ReadRegString(Root: HKEY; const Key, Name: String): string;
var Handle: HKey;
    BufSize: Integer;
    DataType: Integer;
begin
  Result := '';
  if RegOpenKeyEx(Root, PChar(Key), 0, KEY_READ, Handle) = ERROR_SUCCESS then
  begin
    DataType := reg_sz;
    BufSize := 1024;
    SetLength(Result, BufSize);
    if RegQueryValueEx(Handle, PChar(Name), nil, @DataType, @Result[1], @BufSize) <> ERROR_SUCCESS then
      BufSize := 0;
    Dec(BufSize); // Skip the trminating #0
    if BufSize < 0 then BufSize := 0;
    SetLength(Result, BufSize);
    RegCloseKey(Handle);
  end;
end;
{$ENDIF}

{$IFDEF LINUX}
// Read (a number of) paths from the environment
procedure GetEnvPaths(Paths: TStringList; const Env: string);
var S: string;
    p: Integer;
 procedure AddPath(S: string);
 begin
   S := Trim(S);
   if S <> '' then Paths.Add(S);
 end;
begin
  S := GetEnv(PChar(Env));
  repeat
    p := Pos(':', S);
    if p <= 0 then
      AddPath(S)
    else begin
      AddPath((Copy(S, 1, p - 1)));
      S := Trim(Copy(S, p + 1, Length(S)));
    end;
  until p <= 0;
end;
{$ENDIF}

procedure AddToPath(var Path: string; Addition: string);
begin
{$IFDEF LINUX}
  if (Length(Path) > 0) and (Path[Length(Path)] <> '/') then Path := Path + '/';
{$ELSE}
  if (Length(Path) > 0) and (Path[Length(Path)] <> '\') then Path := Path + '\';
{$ENDIF}
  Path := Path + Addition;
end;

procedure OCILog(const S: string);
begin
  if InitOCILog <> '' then InitOCILog := InitOCILog + #13#10;
  InitOCILog := InitOCILog + S;
end;

// Search for Oracle Home
{$IFNDEF LINUX}
function FindHomeKey: string;
var HomeName, HomeKey, ForcedHome, HomeDir, FoundHomeName: string;
    i, HomeIndex, HomePathPos, p: Integer;
    Path: string;
begin
  if OracleHomeKey <> '' then
  begin
    Result := OracleHomeKey;
    Exit;
  end;
  BuildOracleHomeList;
  // Check if there are multiple oracle homes
  if HomeList.Count = 0 then
    Result := 'SOFTWARE\ORACLE'
  else begin
    ForcedHome := AnsiUpperCase(GetParamString('ORACLEHOME'));
    if ForcedHome = '' then ForcedHome := OracleHomeName;
    DefaultHomeName := False;
    FoundHomeName := '';
    HomeKey := '';
    Result := '';
    SetLength(Path, 1000);
    SetLength(Path, GetEnvironmentVariable(PChar('PATH'), PChar(Path), Length(Path)));
    Path := AnsiUpperCase(Path);
    HomePathPos := Length(Path);
    for i := 0 to HomeList.Count - 1 do
    begin
      HomeIndex := Integer(HomeList.Objects[i]);
      HomeName  := HomeList[i];
      HomeKey   := 'SOFTWARE\ORACLE\HOME' + IntToStr(HomeIndex);
      if AnsiUpperCase(HomeName) = AnsiUpperCase(ForcedHome) then
      begin
        Result := HomeKey;
        FoundHomeName := HomeName;
        Break;
      end;
      // Determine position in Path environment
      HomeDir := ReadRegString(HKEY_LOCAL_MACHINE, HomeKey, 'ORACLE_HOME');
      if HomeDir <> '' then
      begin
        AddToPath(HomeDir, 'bin');
        p := Pos(AnsiUpperCase(HomeDir), Path);
        if (p > 0) and (p < HomePathPos) then
        begin
          Result := HomeKey;
          FoundHomeName := HomeName;
          HomePathPos := p;
        end;
      end;
    end;
    // If none detected as preference, use the last one
    if Result = '' then
    begin
      Result := HomeKey;
      FoundHomeName := HomeName;
    end;
    if OracleHomeName = '' then
    begin
      DefaultHomeName := True;
      OracleHomeName  := FoundHomeName;
    end;
  end;
  if ReadRegString(HKEY_LOCAL_MACHINE, Result, 'ORACLE_HOME') = '' then Result := 'SOFTWARE\ORACLE';
  OracleHomeKey := Result;
end;
{$ENDIF}

procedure FindAliases(Path: string);
var T: TStrings;
    i, l, p, Level: Integer;
    F, S: string;
    FMode: Boolean;
begin
  // Remove redundant slashes
  while Copy(Path, 1, 3) = '\\\' do Delete(Path, 1, 1);
  FMode := False;
  T := TStringList.Create;
  try
    if FileExists(Path) then T.LoadFromFile(Path);
    Level := 0;
    for i := 0 to T.Count - 1 do
    begin
      S := '';
      // Clean the line up a bit
      for l := 1 to Length(T[i]) do
      begin
        if T[i][1] = '#' then Break;
        if T[i][l] = '(' then inc(Level);
        if (Level = 0) and (Ord(T[i][l]) >= 32) then S := S + T[i][l];
        if T[i][l] = ')' then dec(Level);
      end;
      S := Trim(S);
      if S <> '' then
      begin
        p := Pos('=', S);
        if p > 0 then
        begin
          F := Trim(Copy(S, p + 1, Length(S)));
          S := Trim(Copy(S, 1, p - 1));
          p := Pos('.WORLD', AnsiUpperCase(S));
          if p > 0 then S := Copy(S, 1, p - 1);
          if Lowercase(S) = 'ifile' then
          begin
            // Look for a file
            FMode := False;
            if F = '' then FMode := True else FindAliases(F)
          end else
            AliasList.Add(S);
        end else begin
          // We were looking for a path
          S := Trim(S);
          if (S <> '') and FMode then
          begin
            FindAliases(S);
            FMode := False;
          end;
        end;
      end;
    end;
  finally
    T.Free;
  end;
end;

function TNSNames: string;
begin
  if OracleTNSNames = '' then
  begin
    {$IFDEF LINUX}
    OracleTNSNames := GetEnv('TNS_ADMIN');
    if OracleTNSNames = '' then
    begin
      OracleTNSNames := OracleHome;
      if OracleTNSNames <> '' then AddToPath(OracleTNSNames, 'network/admin');
    end;
    if OracleTNSNames <> '' then
    try
    {$ELSE}
    FindHomeKey;
    // Initialize the dll, needed for OCI80 boolean
    InitOCI;
    try
      // Get TNS_ADMIN from the environment variable
      SetLength(OracleTNSNames, 1000);
      SetLength(OracleTNSNames, GetEnvironmentVariable(PChar('TNS_ADMIN'),
                PChar(OracleTNSNames), Length(OracleTNSNames)));
      // Get TNS_ADMIN from the registry
      if OracleTNSNames = '' then
        OracleTNSNames := ReadRegString(HKEY_LOCAL_MACHINE, OracleHomeKey, 'TNS_ADMIN');
      // Retry from SOFTWARE\ORACLE if necessary
      if (OracleTNSNames = '') and (StrIComp(PChar(OracleHomeKey), 'SOFTWARE\ORACLE') <> 0) then
        OracleTNSNames := ReadRegString(HKEY_LOCAL_MACHINE, 'SOFTWARE\ORACLE', 'TNS_ADMIN');
      if OracleTNSNames = '' then
      begin
        // Check Oracle8 registry entries
        if OCI80Detected then
          OracleTNSNames := ReadRegString(HKEY_LOCAL_MACHINE, OracleHomeKey, 'NET80')
        else
          OracleTNSNames := ReadRegString(HKEY_LOCAL_MACHINE, OracleHomeKey, 'NET20');
        if OracleTNSNames <> '' then
          AddToPath(OracleTNSNames, 'Admin')
        else begin
          // If TNS_ADMIN, NET80/NET20 not set, use default path
          OracleTNSNames := ReadRegString(HKEY_LOCAL_MACHINE, OracleHomeKey, 'ORACLE_HOME');
          if OCI80Detected and not OCI81Detected then
            AddToPath(OracleTNSNames, 'Net80\Admin')
          else
            AddToPath(OracleTNSNames, 'Network\Admin');
        end;
      end;
    {$ENDIF}
      AddToPath(OracleTNSNames, 'tnsnames.ora');
    except;
      // If something goes wrong, let it go wrong silent
    end;
  end;
  Result := OracleTNSNames;
end;

// Fill OracleAliasList with available database aliases
procedure BuildOracleAliasList;
begin
  if AliasList = nil then
  begin
    // Create if necessary
    AliasList := TStringList.Create;
    AliasList.Sorted := True;
    AliasList.Duplicates := dupIgnore;
  end;
  AliasList.Clear;
  FindAliases(TNSNames);
end;

// Public function to build and return the OracleAliasList
function OracleAliasList: TStringList;
begin
  try
    if AliasList = nil then BuildOracleAliasList;
  except
  end;
  Result := AliasList;
end;

// Fill OracleHomeList with available Oracle Homes
procedure BuildOracleHomeList;
{$IFNDEF LINUX}
var HomeCountString, HomeKey, HomeName: string;
    HomeIndex, HomeCount, Error: Integer;
{$ENDIF}
begin
  if HomeList = nil then
  begin
    // Create if necessary
    HomeList := TStringList.Create;
    HomeList.Sorted := True;
//  HomeList.Duplicates := dupIgnore;
  end;
  HomeList.Clear;
{$IFNDEF LINUX}
  HomeCountString := ReadRegString(HKEY_LOCAL_MACHINE, 'SOFTWARE\ORACLE\ALL_HOMES', 'HOME_COUNTER');
  Val(HomeCountString, HomeCount, Error);
  if Error = 0 then
  begin
    if HomeCount <= 1 then HomeCount := 1;  // HomeCount < 1? Try it anyway
    for HomeIndex := 0 to HomeCount - 1 do
    begin
      HomeKey := 'SOFTWARE\ORACLE\HOME' + IntToStr(HomeIndex);
      HomeName := ReadRegString(HKEY_LOCAL_MACHINE, HomeKey, 'ORACLE_HOME_NAME');
      if HomeName <> '' then HomeList.AddObject(HomeName, TObject(HomeIndex));
    end;
  end;
{$ENDIF}
end;

// Public function to build and return the OracleHomeList
function OracleHomeList: TStringList;
begin
  try
    if HomeList = nil then BuildOracleHomeList;
  except
  end;
  Result := HomeList;
end;

function OracleHome: string;
{$IFDEF LINUX}
var Paths: TStringList;
{$ENDIF}
begin
  {$IFDEF LINUX}
  Paths := TStringList.Create;
  GetEnvPaths(Paths, 'ORACLE_HOME');
  if Paths.Count > 0 then Result := Paths[0] else Result := '';
  Paths.Free;
  {$ELSE}
  Result := ReadRegString(HKEY_LOCAL_MACHINE, FindHomeKey, 'ORACLE_HOME');
  {$ENDIF}
end;

// DLL functions

function GetOCIDLLList(DoLog: Boolean): TStringList;
{$IFDEF LINUX}
var Path: String;
    i: Integer;
    Paths: TStringList;
{$ELSE}
var Path: String;
    Error, Code, Offset: Integer;
    fv: Double;
    DTA: TSearchRec;
{$ENDIF}
begin
  OracleHomeDir := OracleHome;
  Result := TStringList.Create;
{$IFDEF LINUX}
  Paths := TStringList.Create;
  GetEnvPaths(Paths, 'ORACLE_HOME');
  for i := 0 to Paths.Count - 1 do
  begin
    Path := Paths[i];
    if DoLog then OCILog('OracleHome: ' + Path);
    AddToPath(Path, 'lib/');
    Paths[i] := Path;
  end;
  GetEnvPaths(Paths, 'LD_LIBRARY_PATH');
  for i := 0 to Paths.Count - 1 do
  begin
    Path := Paths[i];
    AddToPath(Path, 'libclntsh.so');
    if FileExists(Path) then
    begin
      Result.Add(Path);
      Break;
    end;
  end;
  Paths.Free;
{$ELSE}
  Path := OracleHomeDir;
  if Path <> '' then
  begin
    AddToPath(Path, 'bin\');
    if FileExists(Path + 'oci.dll') then
    begin
      if DoLog then OCILog('Found: ' + 'oci.dll');
      Result.Add(Path + 'oci.dll');
    end;
    Error := FindFirst(Path + 'ora*.dll', faAnyFile,DTA);
    while Error = 0 do
    begin
      Offset := 4;
      if UpperCase(Copy(DTA.Name,4,2)) = 'NT' then inc(Offset, 2);
      val('0.' + Copy(DTA.Name, Offset, Pos('.', DTA.Name) - Offset), fv, Code);
      if (Code = 0) and (fv > 0) then
      begin
        if DoLog then OCILog('Found: ' + DTA.Name);
        Result.Add(Path + DTA.Name);
      end;
      Error := FindNext(DTA);
    end;
    FindClose(DTA);
  end;
{$ENDIF}
end;

// Search for OCI DLL
function FindOCIDLL: Integer;
var List: TStringList;
    DLL: String;
    Excluded, UseAllDLLs: Boolean;
{$IFNDEF LINUX}
    i, Code, Offset: Integer;
    fv, fvMax: Double;
{$ENDIF}
begin
  DefaultOCIDLL := False;
  ExcludedOCIDLLs := UpperCase(ExcludedOCIDLLs);
  Result := dllOK;
  DLL := GetParamString('OCIDLL');           // First check parameter
  if DLL <> '' then OCIDLL := DLL;
  if OCIDLL <> '' then
  begin
    OCILog('OCIDLL forced to ' + OCIDLL);
    Exit;
  end;
{$IFDEF LINUX}
  List := GetOCIDLLList(True);
  if List.Count > 0 then OCIDLL := List[0];
  if OCIDLL = '' then Result := dllNoFile;
  if OCIDLL <> '' then OCILog('Using: ' + OCIDLL);
{$ELSE}
  OracleHomeDir := OracleHome;
  OCILog('OracleHomeKey: ' + OracleHomeKey);
  OCILog('OracleHomeDir: ' + OracleHomeDir);
  List := GetOCIDLLList(True);
  if OracleHomeDir = '' then
    Result := dllNoRegistry
  else begin
    fvMax := 0.0;
    UseAllDLLs := True;
    repeat
      UseAllDLLs := not UseAllDLLs;
      for i := 0 to List.Count - 1 do
      begin
        DLL := ExtractFilename(List[i]);
        Excluded := (not UseAllDLLs) and (Pos('"' + UpperCase(DLL) +'"', ExcludedOCIDLLs) > 0);
        if not Excluded then
        begin
          if UpperCase(DLL) = 'OCI.DLL' then
          begin
            OCIDLL := List[i];
            DefaultOCIDLL := True;
            Break;
          end;
          Offset := 4;
          if UpperCase(Copy(DLL,4,2)) = 'NT' then inc(Offset, 2);
          val('0.' + Copy(DLL, Offset, Pos('.', DLL) - Offset), fv, Code);
          if (Code = 0) and (fv > fvMax) then
          begin
            fvMax  := fv;
            OCIDLL := List[i];
            Result := dllOK;
            DefaultOCIDLL := True;
          end;
        end;
      end;
    until UseAllDLLs or (OCIDLL <> '');
    if OCIDLL = '' then Result := dllNoFile;
    if OCIDLL <> '' then OCILog('Using: ' + OCIDLL);
  end;
{$ENDIF}
  List.Free;
end;

// Function that indicates if OCI DLL is loaded
function DLLLoaded:Boolean;
begin
  {$IFDEF LINUX}
  Result := HDLL <> nil;
  {$ELSE}
  Result := HDLL >= hInstance_Error;
  {$ENDIF}
end;

// Get the address of a procedure in the DLL
procedure GetProc(Handle: THandle; var OK:Boolean; var Ad:pointer; const Name:String);
begin
  if not OK then
    Ad := nil
  else begin
    {$IFDEF LINUX}
    Ad := dlsym(Handle, PChar(Name));
    {$ELSE}
    Ad := GetProcAddress(Handle, PChar(Name));
    {$ENDIF}
    if Ad = nil then OK := False;
  end;
end;

// If necessary, switch back to original default path
procedure ResetOCIPath;
begin
  if OriginalPath <> '' then
  begin
    if GetCurrentDir = ExtractFileDir(OCIDLL) then SetCurrentDir(OriginalPath);
    OriginalPath := '';
  end;
end;

// Initialize OCI DLL
function DLLInit:Integer;
begin
  Result := dllOK;
  if DLLLoaded then Exit;          // exit if already loaded
  InitOCILog := '';
  Result := FindOCIDLL;
  if Result <> dllOK then Exit;
{$IFDEF LINUX}
  HDLL := dlopen(PChar(OCIDLL), RTLD_NOW);
{$ELSE}
  HDLL := LoadLibrary(PChar(OCIDLL));
  if not DLLLoaded then
  begin
    // Try again with Oracles's bin directory as default path
    OriginalPath := GetCurrentDir;
    SetCurrentDir(ExtractFileDir(OCIDLL));
    HDLL := LoadLibrary(PChar(OCIDLL));
  end;
{$ENDIF}
  if not DLLLoaded then
  begin
    OCILog('LoadLibrary(' + OCIDLL + ') returned ' + IntToStr(Integer(HDLL)));
    {$IFDEF LINUX} OCILog(dlerror); {$ENDIF}
    Result := dllNoFile
  end else begin
    // OCI 7 functions
    OCI70 := True;
    GetProc(HDLL, OCI70, @orlon, 'orlon');
    GetProc(HDLL, OCI70, @ologof, 'ologof');
    GetProc(HDLL, OCI70, @oerhms, 'oerhms');
    GetProc(HDLL, OCI70, @oopen, 'oopen');
    GetProc(HDLL, OCI70, @oparse, 'oparse');
    GetProc(HDLL, OCI70, @odefin, 'odefin');
    GetProc(HDLL, OCI70, @oexfet, 'oexfet');
    GetProc(HDLL, OCI70, @ofen, 'ofen');
    GetProc(HDLL, OCI70, @oflng, 'oflng');
    GetProc(HDLL, OCI70, @oclose, 'oclose');
    GetProc(HDLL, OCI70, @odescr, 'odescr');
    GetProc(HDLL, OCI70, @oexec, 'oexec');
    GetProc(HDLL, OCI70, @oexn, 'oexn');
    GetProc(HDLL, OCI70, @ocom, 'ocom');
    GetProc(HDLL, OCI70, @orol, 'orol');
    GetProc(HDLL, OCI70, @ocon, 'ocon');
    GetProc(HDLL, OCI70, @ocof, 'ocof');
    GetProc(HDLL, OCI70, @obreak, 'obreak');
    GetProc(HDLL, OCI70, @obndra, 'obndra');
    GetProc(HDLL, OCI70, @odessp, 'odessp');
    OCI72 := OCI70;
    GetProc(HDLL, OCI72, @olog, 'olog');
    OCI73 := OCI72;
    GetProc(HDLL, OCI73, @opinit, 'opinit');
    // OCI 8 Relational functions
    OCI80 := not ForceOCI7;
    OCI80Detected := True;
    GetProc(HDLL, OCI80Detected, @OCIInitialize, 'OCIInitialize');
    GetProc(HDLL, OCI80Detected, @OCIHandleAlloc, 'OCIHandleAlloc');
    GetProc(HDLL, OCI80Detected, @OCIHandleFree, 'OCIHandleFree');
    GetProc(HDLL, OCI80Detected, @OCIDescriptorAlloc, 'OCIDescriptorAlloc');
    GetProc(HDLL, OCI80Detected, @OCIDescriptorFree, 'OCIDescriptorFree');
    GetProc(HDLL, OCI80Detected, @OCIEnvInit, 'OCIEnvInit');
    GetProc(HDLL, OCI80Detected, @OCIServerAttach, 'OCIServerAttach');
    GetProc(HDLL, OCI80Detected, @OCIServerDetach, 'OCIServerDetach');
    GetProc(HDLL, OCI80Detected, @OCISessionBegin, 'OCISessionBegin');
    GetProc(HDLL, OCI80Detected, @OCISessionEnd, 'OCISessionEnd');
    GetProc(HDLL, OCI80Detected, @OCILogon, 'OCILogon');
    GetProc(HDLL, OCI80Detected, @OCILogoff, 'OCILogoff');
    GetProc(HDLL, OCI80Detected, @OCIPasswordChange, 'OCIPasswordChange');
    GetProc(HDLL, OCI80Detected, @OCIStmtPrepare, 'OCIStmtPrepare');
    GetProc(HDLL, OCI80Detected, @OCIBindByPos, 'OCIBindByPos');
    GetProc(HDLL, OCI80Detected, @OCIBindByName, 'OCIBindByName');
    GetProc(HDLL, OCI80Detected, @OCIBindObject, 'OCIBindObject');
    GetProc(HDLL, OCI80Detected, @OCIBindDynamic, 'OCIBindDynamic');
    GetProc(HDLL, OCI80Detected, @OCIBindArrayOfStruct, 'OCIBindArrayOfStruct');
    GetProc(HDLL, OCI80Detected, @OCIStmtGetPieceInfo, 'OCIStmtGetPieceInfo');
    GetProc(HDLL, OCI80Detected, @OCIStmtSetPieceInfo, 'OCIStmtSetPieceInfo');
    GetProc(HDLL, OCI80Detected, @OCIStmtExecute, 'OCIStmtExecute');
    GetProc(HDLL, OCI80Detected, @OCIDefineByPos, 'OCIDefineByPos');
    GetProc(HDLL, OCI80Detected, @OCIDefineObject, 'OCIDefineObject');
    GetProc(HDLL, OCI80Detected, @OCIDefineDynamic, 'OCIDefineDynamic');
    GetProc(HDLL, OCI80Detected, @OCIDefineArrayOfStruct, 'OCIDefineArrayOfStruct');
    GetProc(HDLL, OCI80Detected, @OCIStmtFetch, 'OCIStmtFetch');
    GetProc(HDLL, OCI80Detected, @OCIStmtGetBindInfo, 'OCIStmtGetBindInfo');
    GetProc(HDLL, OCI80Detected, @OCIDescribeAny, 'OCIDescribeAny');
    GetProc(HDLL, OCI80Detected, @OCIParamGet, 'OCIParamGet');
    GetProc(HDLL, OCI80Detected, @OCIParamSet, 'OCIParamSet');
    GetProc(HDLL, OCI80Detected, @OCITransStart, 'OCITransStart');
    GetProc(HDLL, OCI80Detected, @OCITransDetach, 'OCITransDetach');
    GetProc(HDLL, OCI80Detected, @OCITransCommit, 'OCITransCommit');
    GetProc(HDLL, OCI80Detected, @OCITransRollback, 'OCITransRollback');
    GetProc(HDLL, OCI80Detected, @OCITransPrepare, 'OCITransPrepare');
    GetProc(HDLL, OCI80Detected, @OCITransForget, 'OCITransForget');
    GetProc(HDLL, OCI80Detected, @OCIErrorGet, 'OCIErrorGet');
    GetProc(HDLL, OCI80Detected, @OCILobAppend, 'OCILobAppend');
    GetProc(HDLL, OCI80Detected, @OCILobAssign, 'OCILobAssign');
    GetProc(HDLL, OCI80Detected, @OCILobCharSetForm, 'OCILobCharSetForm');
    GetProc(HDLL, OCI80Detected, @OCILobCharSetId, 'OCILobCharSetId');
    GetProc(HDLL, OCI80Detected, @OCILobCopy, 'OCILobCopy');
    GetProc(HDLL, OCI80Detected, @OCILobDisableBuffering, 'OCILobDisableBuffering');
    GetProc(HDLL, OCI80Detected, @OCILobEnableBuffering, 'OCILobEnableBuffering');
    GetProc(HDLL, OCI80Detected, @OCILobErase, 'OCILobErase');
    GetProc(HDLL, OCI80Detected, @OCILobFileClose, 'OCILobFileClose');
    GetProc(HDLL, OCI80Detected, @OCILobFileCloseAll, 'OCILobFileCloseAll');
    GetProc(HDLL, OCI80Detected, @OCILobFileExists, 'OCILobFileExists');
    GetProc(HDLL, OCI80Detected, @OCILobFileGetName, 'OCILobFileGetName');
    GetProc(HDLL, OCI80Detected, @OCILobFileIsOpen, 'OCILobFileIsOpen');
    GetProc(HDLL, OCI80Detected, @OCILobFileOpen, 'OCILobFileOpen');
    GetProc(HDLL, OCI80Detected, @OCILobFileSetName, 'OCILobFileSetName');
    GetProc(HDLL, OCI80Detected, @OCILobFlushBuffer, 'OCILobFlushBuffer');
    GetProc(HDLL, OCI80Detected, @OCILobGetLength, 'OCILobGetLength');
    GetProc(HDLL, OCI80Detected, @OCILobIsEqual, 'OCILobIsEqual');
    GetProc(HDLL, OCI80Detected, @OCILobLoadFromFile, 'OCILobLoadFromFile');
    GetProc(HDLL, OCI80Detected, @OCILobLocatorIsInit, 'OCILobLocatorIsInit');
    GetProc(HDLL, OCI80Detected, @OCILobRead, 'OCILobRead');
    GetProc(HDLL, OCI80Detected, @OCILobTrim, 'OCILobTrim');
    GetProc(HDLL, OCI80Detected, @OCILobWrite, 'OCILobWrite');
    GetProc(HDLL, OCI80Detected, @OCIBreak, 'OCIBreak');
    GetProc(HDLL, OCI80Detected, @OCIServerVersion, 'OCIServerVersion');
    GetProc(HDLL, OCI80Detected, @OCIAttrGet, 'OCIAttrGet');
    GetProc(HDLL, OCI80Detected, @OCIAttrSet, 'OCIAttrSet');
    GetProc(HDLL, OCI80Detected, @OCISvcCtxToLda, 'OCISvcCtxToLda');
    GetProc(HDLL, OCI80Detected, @OCILdaToSvcCtx, 'OCILdaToSvcCtx');
    GetProc(HDLL, OCI80Detected, @OCIResultSetToStmt, 'OCIResultSetToStmt');
    // OCI 8 Object navigational functions
//    GetProc(HDLL, OCI80Detected, @OCICacheUnpin, 'OCICacheUnpin');
//    GetProc(HDLL, OCI80Detected, @OCICacheFree, 'OCICacheFree');
//    GetProc(HDLL, OCI80Detected, @OCICacheUnmark, 'OCICacheUnmark');
    GetProc(HDLL, OCI80Detected, @OCIObjectUnpin, 'OCIObjectUnpin');
//    GetProc(HDLL, OCI80Detected, @OCIObjectPinCountReset, 'OCIObjectPinCountReset');
    GetProc(HDLL, OCI80Detected, @OCIObjectLock, 'OCIObjectLock');
    GetProc(HDLL, OCI80Detected, @OCIObjectMarkUpdate, 'OCIObjectMarkUpdate');
    GetProc(HDLL, OCI80Detected, @OCIObjectUnmark, 'OCIObjectUnmark');
//    GetProc(HDLL, OCI80Detected, @OCIObjectUnmarkByRef, 'OCIObjectUnmarkByRef');
//    GetProc(HDLL, OCI80Detected, @OCIObjectMarkDeleteByRef, 'OCIObjectMarkDeleteByRef');
    GetProc(HDLL, OCI80Detected, @OCIObjectMarkDelete, 'OCIObjectMarkDelete');
    GetProc(HDLL, OCI80Detected, @OCIObjectFlush, 'OCIObjectFlush');
    GetProc(HDLL, OCI80Detected, @OCIObjectCopy, 'OCIObjectCopy');
    GetProc(HDLL, OCI80Detected, @OCIObjectGetTypeRef, 'OCIObjectGetTypeRef');
    GetProc(HDLL, OCI80Detected, @OCIObjectGetObjectRef, 'OCIObjectGetObjectRef');
    GetProc(HDLL, OCI80Detected, @OCIObjectGetInd, 'OCIObjectGetInd');
    GetProc(HDLL, OCI80Detected, @OCIObjectExists, 'OCIObjectExists');
    GetProc(HDLL, OCI80Detected, @OCIObjectIsLocked, 'OCIObjectIsLocked');
    GetProc(HDLL, OCI80Detected, @OCIObjectIsDirty, 'OCIObjectIsDirty');
    GetProc(HDLL, OCI80Detected, @OCIObjectRefresh, 'OCIObjectRefresh');
    GetProc(HDLL, OCI80Detected, @OCIObjectPinTable, 'OCIObjectPinTable');
    GetProc(HDLL, OCI80Detected, @OCIObjectNew, 'OCIObjectNew');
    GetProc(HDLL, OCI80Detected, @OCIObjectPin, 'OCIObjectPin');
    GetProc(HDLL, OCI80Detected, @OCIObjectFree, 'OCIObjectFree');
    GetProc(HDLL, OCI80Detected, @OCITypeByName, 'OCITypeByName');  // OBSOLETE ?????
//    GetProc(HDLL, OCI80Detected, @OCIObjectArrayPin, 'OCIObjectArrayPin');
//    GetProc(HDLL, OCI80Detected, @OCICacheFlush, 'OCICacheFlush');
//    GetProc(HDLL, OCI80Detected, @OCICacheRefresh, 'OCICacheRefresh');
    GetProc(HDLL, OCI80Detected, @OCIObjectGetAttr, 'OCIObjectGetAttr');
    GetProc(HDLL, OCI80Detected, @OCIObjectSetAttr, 'OCIObjectSetAttr');
    GetProc(HDLL, OCI80Detected, @OCIStringPtr, 'OCIStringPtr');
    GetProc(HDLL, OCI80Detected, @OCIStringAssignText, 'OCIStringAssignText');
    GetProc(HDLL, OCI80Detected, @OCINumberFromInt, 'OCINumberFromInt');
    GetProc(HDLL, OCI80Detected, @OCINumberToInt, 'OCINumberToInt');
    GetProc(HDLL, OCI80Detected, @OCINumberFromReal, 'OCINumberFromReal');
    GetProc(HDLL, OCI80Detected, @OCINumberToReal, 'OCINumberToReal');
    GetProc(HDLL, OCI80Detected, @OCINumberToText, 'OCINumberToText');
    GetProc(HDLL, OCI80Detected, @OCINumberFromText, 'OCINumberFromText');
    GetProc(HDLL, OCI80Detected, @OCIRawPtr, 'OCIRawPtr');
    GetProc(HDLL, OCI80Detected, @OCIRawSize, 'OCIRawSize');
    GetProc(HDLL, OCI80Detected, @OCIRawAssignBytes, 'OCIRawAssignBytes');
    GetProc(HDLL, OCI80Detected, @OCIRefAssign, 'OCIRefAssign');
    GetProc(HDLL, OCI80Detected, @OCIRefClear, 'OCIRefClear');
    GetProc(HDLL, OCI80Detected, @OCIRefIsNull, 'OCIRefIsNull');
    GetProc(HDLL, OCI80Detected, @OCIRefFromHex, 'OCIRefFromHex');
    GetProc(HDLL, OCI80Detected, @OCIRefHexSize, 'OCIRefHexSize');
    GetProc(HDLL, OCI80Detected, @OCIRefToHex, 'OCIRefToHex');
    GetProc(HDLL, OCI80Detected, @OCICollSize, 'OCICollSize');
    GetProc(HDLL, OCI80Detected, @OCICollTrim, 'OCICollTrim');
    GetProc(HDLL, OCI80Detected, @OCICollMax, 'OCICollMax');
    GetProc(HDLL, OCI80Detected, @OCICollGetElem, 'OCICollGetElem');
    GetProc(HDLL, OCI80Detected, @OCICollAssignElem, 'OCICollAssignElem');
    GetProc(HDLL, OCI80Detected, @OCICollAppend, 'OCICollAppend');
    GetProc(HDLL, OCI80Detected, @OCITableDelete, 'OCITableDelete');
    GetProc(HDLL, OCI80Detected, @OCICacheFlush, 'OCICacheFlush');
    if not OCI80Detected then OCI80 := False;
    // External Procedure functions (OCI 805)
    ExtProcDetected := OCI80Detected;
    GetProc(HDLL, ExtProcDetected, @OCIExtProcGetEnv, 'ociepgoe');
    GetProc(HDLL, ExtProcDetected, @OCIExtProcRaiseExcpWithMsg, 'ociepmsg');
    // New OCI 8.1 functions
    OCI81Detected := OCI80Detected;
    OCI81 := OCI81Detected;
    GetProc(HDLL, OCI81Detected, @OCIEnvCreate, 'OCIEnvCreate');
    GetProc(HDLL, OCI81Detected, @OCIDirPathAbort, 'OCIDirPathAbort');
    GetProc(HDLL, OCI81Detected, @OCIDirPathFinish, 'OCIDirPathFinish');
    GetProc(HDLL, OCI81Detected, @OCIDirPathPrepare, 'OCIDirPathPrepare');
    GetProc(HDLL, OCI81Detected, @OCIDirPathColArrayEntrySet, 'OCIDirPathColArrayEntrySet');
    GetProc(HDLL, OCI81Detected, @OCIDirPathColArrayToStream, 'OCIDirPathColArrayToStream');
    GetProc(HDLL, OCI81Detected, @OCIDirPathLoadStream, 'OCIDirPathLoadStream');
    GetProc(HDLL, OCI81Detected, @OCIDirPathStreamReset, 'OCIDirPathStreamReset');
    if not OCI81Detected then OCI81 := False;
    if not (OCI70 or OCI80) then Result := dllMismatch;
  end;
end;

// Initialize the Oracle Call Interface
procedure InitOCI;
var Error: Integer;
    Msg: string;
begin
  OCISection.Enter;
  try
    if not DLLLoaded then
    begin
      Error := DLLInit;
      Msg := '';
      case Error of
      dllNoRegistry : Msg := 'SQL*Net not properly installed';
          dllNoFile : if OCIDLL <> '' then
                        Msg := 'Could not load "' + OCIDLL + '"'
                      else
                        Msg := 'Could not locate OCI dll';
        dllMismatch : Msg := 'Could not initialize "' + OCIDLL + '"';
      end;
      if Error <> dllOK then
      begin
        DLLExit;
        raise Exception.Create('Initialization error'#13#10 + Msg + #13#10#13#10 + InitOCILog);
      end;
      // Never use OCI7's threadsafe mode, you can't use obreak and we do our own mutexing
      if OCI73 and (not OCI80) then opinit(0);
      // We do use OCI8's threadsafe mode though
      if OCI80 and not OCI81 then OCIInitialize(OCI_OBJECT or OCI_THREADED, nil, nil, nil, nil);
    end;
  finally
    OCISection.Leave;
  end;
end;

function OCIVersion: string;
begin
  if OCI81 then Result := 'Version 8.1'
  else
    if OCI80 then Result := 'Version 8.0'
  else
    if OCI73 then Result := 'Version 7.3'
  else
    if OCI72 then Result := 'Version 7.2'
  else
    if OCI70 then Result := 'Version 7.0'
  else
    Result := 'not initialized';
end;

// Free OCI DLL
procedure DLLExit;
begin
{$IFDEF LINUX}
{ Leads to segmentation violation, GLIBC bug?
  if MTSLoaded then dlclose(HMTS);
  if DLLLoaded then dlclose(HDLL);}
  HMTS := nil;
  HDLL := nil;
{$ELSE}
  if MTSLoaded then FreeLibrary(HMTS);
  if DLLLoaded then FreeLibrary(HDLL);
  HMTS := hInstance_Error - 1;
  HDLL := hInstance_Error - 1;
{$ENDIF}
  if DefaultOCIDLL then OCIDLL := '';
  if DefaultHomeName then OracleHomeName := '';
  OCI70 := False;
  OCI72 := False;
  OCI73 := False;
  OCI80 := False;
  OCI81 := False;
  OCI80Detected   := False;
  OCI81Detected   := False;
  ExtProcDetected := False;
  OracleHomeKey   := '';
  OracleHomeDir   := '';
  MTSDetected     := False;
  MTSDLL          := '';
end;

function MTSLoaded: Boolean;
begin
{$IFDEF LINUX}
  Result := HMTS <> nil;
{$ELSE}
  Result := HMTS >= hInstance_Error;
{$ENDIF}
end;

procedure InitMTS;
var Msg: string;
begin
  OCISection.Enter;
  try
    if not MTSLoaded then
    begin
      MTSDetected := False;
      InitOCI;
      Msg := '';
      if MTSDLL = '' then MTSDLL := ExtractFilePath(OCIDLL) + 'oramts.dll';
      if not FileExists(MTSDLL) then Msg := 'Could not locate "' + MTSDLL + '"';
      if Msg = '' then
      begin
        {$IFDEF LINUX}
        HMTS := dlopen(PChar(MTSDLL), RTLD_NOW);
        {$ELSE}
        HMTS := LoadLibrary(PChar(MTSDLL));
        {$ENDIF}
        if not MTSLoaded then Msg := 'Could not load "' + MTSDLL + '"';
        if (Msg = '') and MTSLoaded then
        begin
          MTSDetected := True;
          GetProc(HMTS, MTSDetected, @OraMTSSvcGet, 'OraMTSSvcGet');
          GetProc(HMTS, MTSDetected, @OraMTSSvcRel, 'OraMTSSvcRel');
          GetProc(HMTS, MTSDetected, @OraMTSSvcEnlist, 'OraMTSSvcEnlist');
          GetProc(HMTS, MTSDetected, @OraMTSSvcEnlistEx, 'OraMTSSvcEnlistEx');
          GetProc(HMTS, MTSDetected, @OraMTSTransTest, 'OraMTSTransTest');
          if not MTSDetected then
          begin
            Msg := 'Could not initialize "' + MTSDLL + '"';
            {$IFDEF LINUX}
            dlClose(HMTS);
            HMTS := nil;
            {$ELSE}
            FreeLibrary(HMTS);
            HMTS := hInstance_Error - 1;
            {$ENDIF}
            MTSDLL := '';
          end;
        end;
      end;
      if Msg <> '' then raise Exception.Create('Initialization error'#13#10 + Msg);
    end;
  finally
    OCISection.Leave;
  end;
end;

initialization
begin
  OCISection := TOracleCriticalSection.Create;
end;

finalization
begin
  if AliasList <> nil then AliasList.Free;
  if HomeList <> nil then HomeList.Free;
  OCISection.Free;
end;

end.

