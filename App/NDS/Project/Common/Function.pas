unit James;  {Delphi �Ƶ{���� , �{���}�o : �P���� (James Chou) jmslgs@ms5.hinet.net 1999.06.25 ��z }

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DB, DBTables, Grids, IniFiles, ShellApi, Winsock, ComObj;

//====================================
// ***** �ƾ�
//====================================
function _Max(I1,I2:integer):integer;
function _Min(I1,I2:integer):integer;
function _MinInteger(s:single):Longint; {�D�j��Y��Ƥ��̤p���}
//====================================
// ***** �ഫ
//====================================
function _Int(R:REAL):LONGINT; {�P DELPHI �� INT() ���P,��Ǧ^ LONGINT , �ӫD REAL ( �p123.0 ) }
function _StrToInt(S:string):LONGINT; {�P DELPHI �� strtoint() ���P,�榡�����Ǧ^ 0 , �Ӥ��O���~�T��}
function _Value(S:string):Extended;
function _StrToFloat(str:string):Extended;
function _StrToPchar(S:string):PChar;
function _LineFeedEnCode(Source:string):string; {�N�_��r���s�X���@���r;����X����}
function _LineFeedDeCode(Source:string):string; {�N�s�X�ᤧ�_��r���ѽX}
procedure _StrToStrings(S:string; T:Tstrings); {�N�t���r�I���r���ഫ��Tstrings,��VCL��}
procedure _StrToStrlist(S:string; var T:TstringList); {�N�t���r�I���r���ഫ��TstringLIST}
procedure _StrListToStr(L:Tstrings; var S:string ); {�NTstringLIST�ഫ���t���r�I���r��( �P_StrToStrlist()�ۤ� ) }
procedure _FileList(Path:string;List:Tstrings); {ex: _FileList('c:',listbox1.items) or _FileList('c:', TStringList) }
//====================================
// ***** �r��B�z
//====================================
function _Space(N:integer):string; {�P Clipper SPACE() }
function _RepeatStr(S: string;N:integer): string;
function _RCopy(S:string ; Index,Count:Integer):string;
function _LTrim(S:string):string;  { �h�r�� ���ť�}
function _RTrim(S:string):string; {�P Clipper RTRIM() �h���r��k��ť�}
function _AllTrim(cText: string): string; { �N�r�ꤤ���k�Ҧ����ťզr���R�� }
function _StrZero(N:Longint;L:integer):string; {�P Clipper STRZERO() �Ʀr�ܤ�r��,�e���ɺ�'0' }
function _ReStr(S:string;L:integer):string; {����r����� }
function _Left(S:string;L:integer):string; {�P Clipper Left() }
function _Right(S:string;L:integer):string; {�P Clipper RIGHT() }
function _CutLeft(S:string;L:integer):string; {��r�ꤤ����Y�z�Ӧr��}
function _CutRight(S:string;L:integer):string; {��r�ꤤ�k��Y�z�Ӧr��,�h�ɦW�ᤧ�����ɦW�ɫܦn��}
function _StrRange(S,R:string):Boolean; {�P�_ S �O�_�b R �϶��r�� �Y 'R1-R2' ��}
function _StrMutiRANGE(S,R:string):Boolean;
function _StrBelong(S,R:string):Boolean; {���\���_StrMutiRANGE()��ƧP�_��i}
function _StrLike(S,R:string):Boolean; {���\����_StrBelong()���,��������j�p�g}
function _CutCharRight(S,C:string;F:integer):string; {��r��S��,�̥k��C�r���k��H�ᤧ�r�� ; F=0 �O�dC , F<>0 �sC �h��}
function _CutAllChar(S:string;C:char):string;  {�N�r�� S ���Y��@�r�� C ���Ʊ���}

function _SegNum(S,Seg:string):integer; {�P�_S�r�ꤤ���h��"�`" ( �Y SEG='' �h�H�r�I���` ) }
function _SubStrNum(S,SUBS:string):integer; {���r�ꤤ�X�{�Y�@�l�r�ꤧ���� }
function _SegStr(S:string;I:integer):string; {���X S �r�ꤤ�� I �`�r��(�H�r�I��)}
function _GetSegStr(S:string ; Seg:string ; N:integer):string; {���o S �r�ꤤ�H Seg �l�r����j���� N �Ӧr��Ϭq , ���q��!! }
function _GetCharSegStr(S:string ; C:Char ; N:integer):string; {���o S �r�ꤤ�H C �r���Ÿ����j���� N �Ӧr��Ϭq}

function _FilterNumStr(S: string): string; { �N�r�ꤤ'�Ʀr�r��'�o�X }
function _StrRightSide(S:string ; L:integer):string; {�r��k�a}
function _StrPos(S,subS:string):integer; {�D �r��subs �X�{�� �r��S ���ĴX��}
function _StrKeyWordValue(S,KeyWord:string ; Len:integer):string; {�� �r��S ���X�{ �r��KeyWord ��s���Len�Ӧr��X��}
procedure _ReplaceStr(var SourceStr:string ; OldStr,NewStr:string);
procedure _ReplacePuncStr(var SourceStr:string); {�N���I�Ÿ��������j�g}
procedure _ReplaceSegStr(var S:string ; Seg:string ; SegIndex:integer ; ReplaceStr:string);
procedure _SegStrValueInc(var S:string ; Seg:string ; SegIndex:integer);
procedure _StrSwap(var S1,S2:string);
//====================================
// ***** TStrins ����
//====================================
function _AddStrings(S1,S2 : Tstrings):Tstrings;
function _SeekItemIndex(S:string;ITM:Tstrings):integer; {���o�Y�@�r��S �b�@ Tstrings (�pComBoBox��ITEMS) ���̱��� INDEX ��}
function _SeekItemText(S:string;ITM:Tstrings):string; {���o�Y�@�r��S �b�@ Tstrings (�pComBoBox��ITEMS) ���̱��񤧦r��}
function _ItemExist(T:TstringList;item:string):Boolean;
function _stringListValue(T:TstringList ; Para:string):string;
procedure _TStringsSaveToINISec(TS:Tstrings ; FileName,SEC:string);
procedure _TstringsLoadFromINISec(TS:Tstrings ; FileName,SEC:string);
procedure _GetLeftStrings(S:Tstrings;L:integer; D:Tstrings);
procedure _GetRightStrings(S:Tstrings;L:integer; D:Tstrings);
procedure _DeleteStrings(S:Tstrings;P:integer);
procedure _InsertStrings(S:Tstrings;P:integer;Str:string);
//====================================
// ***** StringGrid ����䴩���
//====================================
function _StrGridText(StrG:TstringGrid):string;
function _StrGridNextCell(S: TstringGrid):Boolean;
function _IsEmptyRow(S: TstringGrid ; R : LongInt):Boolean;
function _IsInitStrGrid(S: TstringGrid ):Boolean;
function _IsLastRow(S: TstringGrid ):Boolean;	{�O�_�b�̫�@�� ROW }
function _StrGridCellRect(Sender: TObject; ACol, ARow: Longint):TRect;
function _StrGridGotoKeyRow(SG:TstringGrid ; key_field:integer ; key:string):Boolean;
procedure _StrGridDelete(S: TstringGrid);
procedure _StrGridInsert(S: TstringGrid);
procedure _StrGridAppend(S: TstringGrid);
procedure _StrGridAdd(S: TstringGrid);
procedure _StrGridDelEmptyRow(S: TstringGrid);
procedure _StrGridInit(S: TstringGrid);   {��l�� stringGrid �u�d�@ ROW }
procedure _StrGridClear(S: TstringGrid);
procedure _StrGridDispLabel(S: TstringGrid ; DP: string);
procedure _StrGridSetText(S: TstringGrid ; T:string);
procedure _StrGridDeleteRow(S: TstringGrid ; R : LongInt);
procedure _StrGridInsertRow(S: TstringGrid ; R : LongInt);
procedure _StrGridClearRow(S: TstringGrid ; R : LongInt);
procedure _StrGridSaveToFile(StrG: TstringGrid ; FileName: string);
procedure _StrGridLoadFromFile(StrG: TstringGrid ; FileName: string);
procedure _StrGridSaveToLines(StrG:TstringGrid ; L:Tstrings ; formatStr:string);
procedure _StrGridLoadFromLines(StrG:TstringGrid ; L:Tstrings ; formatStr:string);
procedure _StrGridSaveToINISec(StrG: TstringGrid ; FileName,Sec: string ; all:integer);
procedure _StrGridLoadFromINISec(StrG: TstringGrid ; FileName,Sec: string);
procedure _StrGridGetINISec(StrG:TstringGrid ; INIFile,Section:string);  {�u�O�C�@�C���e���col,�ΨӰO�ѼƳ]�w��}
procedure _StrGridWriteINISec(StrG:TstringGrid ; INIFile,Section:string); {�u�O�C�@�C���e���col,�ΨӰO�ѼƳ]�w��}
procedure _StrGridTextWrite(StrG:TstringGrid ; Text:string);
//====================================
// ***** INI �� & counter �B�z
//====================================
function _GetINIStr(FileName,Section,KeyName:string):string;
function _WriteINIStr(FileName,Section,KeyName,Parameter:string):Boolean;
function _ReadWriteINIStr(FILE_NAME,SEC,KEYNAME,DefauStr:string):string;
function _GetINISecKeyNum(FILE_NAME,SEC:string):integer;
function _LaunchCounter:integer;
function _CGICounter(CounterName:string):integer;

function _ReadPassWord(IniFile,Section,VarName,DefauVal:string):string; {INI KeyWord �r��ѽX}
function _ReadStr(IniFile,Section,VarName,DefauVal:string):string;
function _WriteStr(IniFile,Section,VarName,VarValue:string):Boolean;
function _ReadInt(IniFile,Section,VarName:string ; DefauVal:integer):integer;
function _WriteInt(IniFile,Section,VarName:string ; VarValue:integer):Boolean;
function _ReadBoolean(IniFile,Section,VarName:string ; DefauVal:Boolean):Boolean;
function _WriteBoolean(IniFile,Section,VarName:string ; VarValue:Boolean):Boolean;

function _QRStr(VarName,DefauVal:string):string; {Quick ReadStr}
function _QWStr(VarName,VarValue:string):Boolean; {Quick WriteStr}
function _QRInt(VarName:string ; DefauVal:integer):integer; {Quick ReadInt}
function _QWInt(VarName:string ; VarValue:integer):Boolean; {Quick WriteInt}

procedure _DelINISection(FILE_NAME,SEC:string);
procedure _GetINISection(FILE_NAME,SEC:string; T:Tstrings);
procedure _GetINISecKeyWord(FILE_NAME,SEC:string; T:Tstrings);
procedure _GetINISectionNames(INIFile:string ; T:Tstrings);
procedure _GetINISectionKeys(INIFile,Section:string ; T:Tstrings ; Val:integer);
procedure _SectionCopy(SourceINI,DestINI,Section:string ; Update:integer);

procedure _TimeHit(EveryDay:integer); {���G�ɶ��O���s�J INI �ɤ� , EveryDay<>0 �ݨC�魫�m}
//====================================
// ***** ��r�ɳB�z
//====================================
function _LoadFromFile(FileName:string;P:integer):string;
function _File(FileName:string):Boolean;
function _FileLineCount(FileName:string):LongInt; {�D�ɮפ���Ʀ��}
procedure _SaveToFile(FileName,S:string);  {�л\}
procedure _AppendToFile(FileName,S:string);
procedure _NewFile(FileName:string);
procedure _DelFile(FileName:string);
//====================================
// ***** ��� DUMP , �ɮ���X
//====================================
procedure _TableToExcel(TB:TDataSet);
procedure _TableDump(TB:TDataSet ; IntField:string);  {Table �ɦL����r��,�i�]�w�k�a���(�ƭ���) }
procedure _TableSelectDump(TB:TDataSet ; SelectField:string); {Table �ɦL����r��,�i������ɦL }
procedure _TableCountDump(TB:TDataSet ; CountField:string); {Table �ɦL����r��(�i�]���`�p��������)}
procedure _StrGridDump(SG:TstringGrid ; DisplayWidth:string);
procedure _StructureDump(TB:TTable);
//====================================
// ***** ��ܲ�,������T
//====================================
function _Question(Q:string):Boolean; { ���D��� ��ܲ� }
function _Input(WhatInput,note,ReadyStr:string):string; { ��J ��ܲ� }
procedure _Show(S:string); { �D�ܹ�ܲ� }
procedure _ShowBoolean(B:Boolean);
procedure _ShowInt(I:LongInt);
procedure _ShowFloat(F:Single);
procedure _SaveLog(mess:string ; replace:integer);
procedure _ViewLog;
//====================================
// ***** CALL �~���{��
//====================================
procedure _Run(ExeStr:string);
procedure _notepad(FileName:string);
procedure _emailto(email_address:string);
procedure _wwwto(url:string);
//====================================
// ***** �s/�ѽX,�P�_,��T���o
//====================================
function _encrypt(S:string):string; {�r��s/�ѽX}
function _AppPath:string;
function _AppIni:string;   {�t���|}
function _AppName:string;  {�t���|}
function _AppExeName:string;  {���t���|,�u���ɦW}
function _ChangeExtName(FN,ExtType:string):string; {�N�ɦW�����O�����ɦW}
function _isEnglishWin:Boolean;
function _isChineseWin:Boolean;
function _isNT : Boolean;
function _myIP:string;
function _IsEmail(email:string):Boolean;
function _AutoCaption(English,Chinese:string):string; {�̷Ӥ��^��������� caption}
function _isLeapYear(yyyy:string):Boolean; {�D�褸�~�r��O�_���|�~}
function _isNullStr(s:string):Boolean;  {�P�_�O�_���Ŧr��,�����ťդ]��Ŧr��}
function _isNullTStrings(TS:Tstrings):Boolean; {�P�_�O�_���Ŧr��զ�,�����ťդ]��Ŧr��}
function _isNumStr(s:string):Boolean;
function _isLetter(s:string):Boolean;
function _isNumKey(key:word):Boolean;
function _isEAN13(ean:string):Boolean;
function _IdCheck(ID: string): Boolean; { �����Ҹ��X�k���ˬd }
function _GetAge(dDate : TDateTime): integer;	{���o�~��}
function _WeekStr(D:TDateTime):string;
function _CountryCode(Tel_Number:string):string;
//====================================
// ***** API , ����
//====================================
function _Resource(GU:string):string; {���o�Ѿl�귽��,�ǤJ�ѼƦ�'GDI'��'USER'}
procedure _CloseOtherAP(Windows_Caption:string);
procedure _Delay(ms : longint);
procedure _Wait(S:Single); { �ɶ�����(59��) �i��J�p�� }
procedure _ForceExec(filename:string);
procedure _keyStep(Key:Word ; option:string);
procedure _Next; {�۰ʱN�ʧ@���� forM ���J�I���ܨ䤺�U�@����}
procedure _UP; {�۰ʱN�ʧ@���� forM ���J�I���ܨ䤺�W�@����}
procedure _ReSetFocus;
procedure _CapLock(bLockIt: Boolean); { ��L�j�p�g��w }
//====================================
// ***** DateTime ����ɶ��B�z
//====================================
function _WhatTime(What:string):string; { ����{�b�ɶ�����,��,�� }
function _WhatDate(What:string):string; { ����{�b������~,��,�� }
function _AddYearMonth(yymm:string ; inc:integer):string; {�D�ǤJ�~�뤧�U�@�~��}
function _DecYearMonth(yymm:string ; dec:integer):string; {�D�ǤJ�~�뤧�e�@�~��}
function _ThisYearMonth(f:integer):string; {�D�{�b�~��}
function _AddTimeStr(TimeStr1,TimeStr2:string):string;  {�ɶ��r��ۥ[ , ��q�T�ɶ��O�ήɥ�}
function _TimeToMinute(TimeStr:string):string; {�ɶ��ন����}
function _TimeToSecond(TimeStr:string):string; {�ɶ��ন����}
function _monthLastDay(yyyymm:string ; onlyday:integer):string; {�D�褸�~�뤧�Ӥ�̫�@�Ѥ���r�� ; onlyday=1 ��u�Ǧ^�Ӥ�ѼƦr��}

function _now:string; {�D�{�b����ɶ�(�褸�~) 1999/07/03 14:10:00}
function _now2:string; {�D�{�b����ɶ�(�褸�~) 1999.07.03 14:10:00}
function _nowstr:string; {�s�ɮװߤ@�W�r�ɨ��W�� 19990703141000}
function _date(i:integer ; c:string):string; {1999/07/03 or 1999.07.03 or 1999-07-03}
function _dateToStr(D:TDatetime ; c:string):string; {1999/07/03 or 1999.07.03 or 1999-07-03}
function _dateToYYYYMMDD(i:integer):string; {19990703}
function _dateToMMDD(i:integer):string; {0703}
//====================================
// ***** ROC DateTime �������ɶ��B�z
//====================================
function _ROCDATE(DD:TDATETIME;P:integer):string; {�ഫ�Y���������0YYMMDD �����r�� }
function _ROCCHECK(DD:string):Boolean;
function _ROCADD(DD:string;N,P:integer):string;
function _ROCDEC(DD1,DD2:string):LONGINT;   {�DDD1-DD2 ����������t}
function _RocNow(T:integer):string;
function _RocReStr(DD:string;T:integer):string; {�������r��(����������T���ˬd)}
function _ROCCheckF(ROC:string;F:integer):Boolean;
function _ROCMonthDay(ROCY,MM:string):integer;
function _ROCLYearCheck(ROCY:string):Boolean;
function _ROCGetSeg(ROC:string;Seg:integer):string; {��������r�ꤤ���~���(�������T���ˬd)}
function _RocToDate(Roc:string):TDateTime;
function _ROCMonthFirstDay(ROCY,M:string;P:integer):string;
function _ROCMonthLastDay(ROCY,M:string;P:integer):string;
//====================================
// ***** �q�T�O�v , �r��B�I�B�z
//====================================
function _BaseNumber(Son,Mon:LongInt):LongInt; {�D�����}
function _str(f:single ; digits:integer):string; {�����T�w�p�Ʀ줧�ƭȦr�� --  ���->�r��  }
function _string(f:string ; digits:integer):string; {�����T�w�p�Ʀ줧�ƭȦr�� --  �r��->�r��  }
function _cost(TimeBase:integer ; BaseRate,ExtenRate:single):single; {�D�q�H�O�v;��ƶǤJ,��ƶǥX}
function _cost2(TimeBase:integer ; BaseRate,ExtenRate:string):string; {�D�q�H�O�v;�r��ǤJ,�r��ǥX(�@��p�Ʈ榡)--> ������� }
function _charge(TimeBase , BaseTime , ExtenTime , BaseRate , ExtenRate:single):single; {�D�q�H�O�v}
//====================================
// ***** EAN_13
//====================================
function _EAN13CheakCode(ean:string):string;
function _toBeEAN13(ean:string):string;
function _serialNumAdd(serial:string):string;
function _serialNumAddEAN13(ean13:string):string;
//====================================
// ***** DATABASE ��Ʈw
//====================================
function _TableFind(TB:TTable ; IndexName:string ; parameter:string):Boolean;  { IndexName='' �ϥέ���� }
function _Count(TB:TDataSet ; CountFields,Digits:string):string; {�h�椧 ���displayText�� (�H�r�I��)�p�ƾ�,}
function _FieldType(FT:TFieldType):string;
function _TableNameExist(AliasName,TableName:string):Boolean;
procedure _TableGroup(TB,TB2:TTable ; MainFieldID,CountFields,Digits:string); {���g����}
procedure _GroupFieldToItem(TB:TTable ; FieldIndex:integer ; Item:Tstrings ; SpaceItem:integer);
procedure _TableEmpty(TB: TTable); {�M��table,�����]�M���ݩ�}
procedure _TableStructureCopy(Source,Dest:TTable ; TableName:string); {�ƻs���c,�L�k�ƻs����}
procedure _TableNew(Source,Dest:TTable ; TableName:string); {�P�W,�i�ƻs���c,�L�k�ƻs����}
//====================================
// ***** DataBase Alias ��Ʈw�O�W�Ѽ�
//====================================
function _AliasDataBase(AliasName:string):string; {�Ǧ^���|�̫�L'\'�r�� ; �Ǧ^DATABASE ��m ----> ���|�θ�Ʈw}
function _AliasDataBasePath(AliasName:string):string; {�Ǧ^ DATABASE ���|��m ----> �������| , �Ǧ^���|�̫�L'\'�r��}
function _Alias(AliasName:string):Boolean;
function _AliasType(AliasName:string):string;
function _AliasDefaultDriver(AliasName:string):string;
function _ReadAliasAppIniStr(AliasName,VarName,DefauVal:string):string;
function _ReadAliasAppIniInt(AliasName,VarName:string ; DefauVal:integer):integer;
procedure _AliasRemove(AliasName:string);
procedure _AliasCreate(AliasName,Path:string ; Driver:integer);
procedure _AliasSetup(AliasName,Path:string ; Driver:integer);
//====================================
// ***** HTML �䴩���
//====================================
function _TableToHTMLContent(TB:TDataSet ; header,footer:string ; fullHTML:integer):string;
function _TableToHTMLEdit(TB:TDataSet ; header,footer:string ; fullHTML:integer ;
                          my_ap:string ; hide_field:string ; new_field:integer):string;
function _StrGridToHTML(SG:TstringGrid ; header,footer:string ; fullHTML:integer ;
                        my_ap:string ; TableWidth:string ; hide_field:string ; key_field:integer ):string ;
function _StrGridToHTMLEdit(SG:TstringGrid ; header,footer:string ; fullHTML:integer ;
                            my_ap:string ; hide_key , hide_field , readonly_field ,
                            new_headfield , new_footfield:string):string;
function _HypLnkText(C,DispText:string):string;
function _HTMLConfirmContent(msg,http:string):string;
function _strToASCII(s:string):string;
function _BigSpace(s:string):string; {�Y���Ŧr��h�Ǧ^�����ť�}
function _CounterToHtml:string;  {���Y�@HTML Table cell ��}
function _HTMLTimeHit(TimeHitStr:string):string;
function _HTMLTimeHit2(TimeHitStr:string):string; //�A�˹�

//====================================
// ***** �䥦
//====================================
procedure _DelTree(dir:String);
function _ToCarryAs(decimal,nToCarry:longint):string;

implementation

//==============================================================================
// ***** �ƾ�
//==============================================================================
function _Max(I1,I2:integer):integer;
begin
   if I1>=I2 then Result:=I1 else
      Result:=I2;
end;
{==============================================================================}
function _Min(I1,I2:integer):integer;
begin
   if I1<=I2 then Result:=I1 else
      Result:=I2;
end;
{==============================================================================}
function _MinInteger(s:single):Longint; {�D�j��Y��Ƥ��̤p���}
var r:Longint;
begin
  r:=_int(s);
  if s>r then r:=r+1;
  Result:=r;
end;
//==============================================================================
// ***** �ഫ
//==============================================================================
function _Int(R:REAL):LONGINT;
{�P DELPHI �� INT() ���P,��Ǧ^ LONGINT , �ӫD REAL ( �p123.0 ) }
var RR:string;
begin
  RR:=FLOATTOSTR(INT(R));
  Result:=strtoint(RR);
end;
//==============================================================================
function _StrToInt(S:string):LONGINT;
{�P DELPHI �� strtoint() ���P,�榡�����Ǧ^ 0 , �Ӥ��O���~�T��}
begin
  Result:=_Int(_Value(S));
end;
//==============================================================================
function _Value(S:string):Extended;
var I:LONGINT;     {88.01.20 �ק�}
    CODE:integer;
    SS:string;
begin
S:=_AllTrim(S);  {�N�r��e��ťեh��}

if _SubStrNum(S,'.')=0 then
   begin
     VAL(S,I,CODE);
     Result:=I;
     exit;
   end
else
   begin
     if _SubStrNum(S,'.')>1 then
	begin
	Result:=0;
	exit;
	end;

     SS:=S;
     Delete(SS,(POS('.',SS)),1);

     Val(SS,I,CODE);

     if I=0 then
	begin
	Result:=0;
	exit;
	end;

    Result:=STRTOFLOAT(S);

   end;
end;
//==============================================================================
function _StrToFloat(str:string):Extended;
var s,ss:string;
    i:integer;
    r:Extended;
begin
  r:=0;
  s:=_AllTrim(str);
  if s='' then       {---------- Gate 1}
     begin
       Result:=r;
       exit;
     end;

  ss:='';
  for i := 1 to Length(s) do
       if (s[i] in ['0'..'9']) or (s[i]='-') or (s[i]='.') then
	  ss := ss + s[i];

  if (_SubStrNum(ss,'-')>1) or (_SubStrNum(ss,'.')>1) then    {---------- Gate 2}
     begin
       Result:=r;
       exit;
     end;

  if (_SubStrNum(ss,'-')=1) and (_Left(ss,1)<>'-') then        {---------- Gate 3}
     begin
       Result:=r;
       exit;
     end;

  if s<>ss then  {---------- Gate 4}
     begin
       Result:=r;
       exit;
     end;

  Result:=strtofloat(ss);

end;
//==============================================================================
function _StrToPchar(S:string):PChar;
VAR
  //P: array[0..255] of Char;
  P: array[0..1000] of Char; {1999.07.08 ��}
begin
  StrPCopy(P, S);
  Result:=STRNEW(P);
end;
//==============================================================================
function _LineFeedEnCode(Source:string):string; {�N�_��r���s�X���@���r;����X����}
var i,skip:integer;
begin
  Result:='';
  skip:=0;
  for i:=1 to length(Source) do
    begin
      if skip<>0 then
         begin
           skip:=skip-1;
           continue;
         end;
      if copy(Source,i,2)=(chr(13)+chr(10)) then
         begin
           Result:=Result+'%0D%0A';
           skip:=1;
         end
      else
         Result:=Result+Source[i];
    end;
end;
//==============================================================================
function _LineFeedDeCode(Source:string):string; {�N�s�X�ᤧ�_��r���ѽX}
var i,skip:integer;
begin
  Result:='';
  skip:=0;
  for i:=1 to length(Source) do
    begin
      if skip<>0 then
         begin
           skip:=skip-1;
           continue;
         end;
      if copy(Source,i,6)='%0D%0A' then
         begin
           Result:=Result+chr(13)+chr(10);
           skip:=5;
         end
      else
         Result:=Result+Source[i];
    end;
end;
//==============================================================================
procedure _StrToStrings(S:string; T:Tstrings);
	 {�N�t���r�I���r���ഫ��TstringLIST}
var P,I : integer;
begin
T.CLEAR;
P:=_SubStrNum(S,',');{�p��ǤJ�r�ꤤ���X�ӳr�I}

if P=0 then
   begin
   T.ADD(S);
   exit;
   end;

for I := 1 to (P+1) do
    begin
      if I=(P+1) then
      T.ADD(S)
    else
      T.ADD(_Left(S,POS(',',S)-1));

      S:=_CUTLeft(S,POS(',',S));
    end;

end;
//==============================================================================
procedure _StrToStrlist(S:string; var T:TstringList);
	 {�N�t���r�I���r���ഫ��TstringLIST}
var P,I : integer;
begin
  T.Clear;
  P:=_SubStrNum(S,',');{�p��ǤJ�r�ꤤ���X�ӳr�I}

  if P=0 then
     begin
       T.Add(S);
       exit;
     end;

  for I := 1 to (P+1) do
    begin
      if I=(P+1) then
        T.Add(S)
      else
        T.Add(_Left(S,Pos(',',S)-1));

      S:=_CutLeft(S,Pos(',',S));

    end;

end;
//==============================================================================
procedure _StrListToStr(L:Tstrings; var S:string ); {�NTstringLIST�ഫ���t���r�I���r��( �P_StrToStrlist()�ۤ� ) }
var I : integer;
begin
  S:='';

  if L.Count=0 then
     begin
       exit;
     end;

  if L.Count=1 then
     begin
       S:=L[0];
       exit;
     end;

  if L.Count>1 then
     begin
       S:=L[0];
       for I := 1 to (L.Count-1) do
	   begin
	     S:=S+','+L[I];
	   end;
     end;

end;
//==============================================================================
procedure _FileList(Path:string;List:Tstrings); {ex: _FileList('c:',listbox1.items) or _FileList('c:', TStringList) }
var p:string;
    r:integer;
    MySearchRec:TSearchRec;
begin
  p:=path;
  if copy(p,length(p),1)<>'\' then p:=p+'\*.*' else p:=p+'*.*';
  list.clear;
  r := FindFirst(p, faAnyFile, MySearchRec);
  while r = 0 do
    begin
      //if MySearchRec.Attr<>faDirectory then listbox1.items.add(MySearchRec.Name);
      if (ExtractFileExt(MySearchRec.Name)<>'') and (MySearchRec.Name<>'.') and (MySearchRec.Name<>'..') then
          list.add(MySearchRec.Name{+';'+ExtractFileExt(MySearchRec.Name)});
          {�Y�ɮ׵L���ɦW�]�|�Q�L�o��}
      r:= FindNext(MySearchRec);
    end;
  FindClose(MySearchRec);
end;
//==============================================================================
// ***** �r��B�z
//==============================================================================
function _Space(N:integer):string; {�P Clipper SPACE() }
var I:integer;
    R:string;
begin
  R:='';
  if N<>0 then
     begin
       for I := 1 to N DO
	   R:=R+' ';
     end;
  Result:=R;
end;
//==============================================================================
function _RepeatStr(S: string;N:integer): string;
var i:integer;
begin
  Result:='';
  if N<=0 then exit;
  for i:=1 to N do
      begin
	Result:=Result+S;
      end;
end;
//==============================================================================
function _RCopy(S:string ; Index,Count:Integer):string;
var R:string;
begin
  R:='';
  if Index>=Length(S) then
     begin
     {SHOWMESSAGE('_RCOPY()��ƿ��~,�I����m�W�X�r�����'); }
     Result:=R;
     Exit;
     end;
  Index:=Length(S)-Index+1;
  R:=Copy(S,Index,Count);

  Result:=R;
end;
//==============================================================================
function _LTrim(S:string):string;  { �h�r�� ���ť�}
var I:integer;
    R:string;
begin
  R:='';
  if (S<>'') and (S<>' ') then
     begin
       for I :=1 to Length(S) do
	   if copy(S,I,1)<>' ' then break;

     R:=COPY(S,I,(Length(S)-I+1));
     if R=' ' then R:='';

     end;
  Result:=R;
end;
//==============================================================================
function _RTrim(S:string):string; {�P Clipper RTRIM() �h���r��k��ť�}
var I:integer;
    R:string;
begin
  R:='';
  if (S<>'') and (S<>' ') then
     begin
       for I :=Length(S) downto 1 do
	   if copy(S,I,1)<>' ' then break;
     R:=COPY(S,1,I);
     if R=' ' then R:='';
     end;
  Result:=R;
end;
//==============================================================================
function _AllTrim(cText: string): string; { �N�r�ꤤ���k�Ҧ����ťզr���R�� }
begin
   Result := _LTrim(_RTrim(cTEXT));
end;
//==============================================================================
function _StrZero(N:Longint;L:integer):string; {�P Clipper STRZERO() �Ʀr�ܤ�r��,�e���ɺ�'0' }
var R:string;
begin
  {���ɶǤJ�ȥ��g�Lstrtoint()�ഫ�|�ܦ��t�Ψ䥦����}
  {�G�Ǧ^�������g�L����ƳB�z��|�ܤ�,���Ostrtoint()�����D}

  if N<0 then
  begin
  showmessage('�ǤJSTRZERO( )��Ƥ���ƭȤp��0 !');
  exit;
  end;

  R:=inttostr(N);
  if Length(R)>L then
     begin
     showmessage('�ǤJ_STRZERO( )��Ƥ���ƪ��׶W�L���w���� ('+inttostr(L)+') !');
     Result:=R;
     exit;
     end;

  while Length(R)<L do
    R:='0'+R;

  Result:=R;
end;
//==============================================================================
function _ReStr(S:string;L:integer):string; {����r�����,��� TABLE �r�����ɫܦn��}
var R:string;				    {�� _RESTR('TEST',6) �Y��N'TEST' �� 'TEST  ' }
begin
  R:=S;
  if Length(R)>L then
     begin
       Result:=COPY(R,1,L);
       exit;
     end;

  while Length(R)<L do
    R:=R+' ';

  Result:=R;
end;
//==============================================================================
function _Left(S:string;L:integer):string; {�P Clipper Left() }
var R:string;
begin
  R:='';

  if (S<>'') and (L<Length(S)) and (L>0) then
     begin
      R:=COPY(S,1,L);
     end;

  if L>=Length(S) then R:=S;

  Result:=R;
end;
//==============================================================================
function _Right(S:string;L:integer):string; {�P Clipper RIGHT() }
var R:string;
begin
  R:='';

  if (S<>'') and (L<Length(S)) and (L>0) then
     begin
      R:=COPY(S,(Length(S)-L+1),L);
     end;

  if L>=Length(S) then R:=S;

  Result:=R;
end;
//==============================================================================
function _CutLeft(S:string;L:integer):string; {��r�ꤤ����Y�z�Ӧr��}
var R:string;				      {�� _CUTRIGHT('TEST',1) �Y�Ǧ^'EST'}
begin
  {R:=S;}
  R:='';
  if (S<>'') and (L<=Length(S)) and (L>0) then
     begin
      R:=_RIGHT(S,(Length(S)-L));
     end;

  if L=0 then R:=S;

  Result:=R;
end;
//==============================================================================
function _CutRight(S:string;L:integer):string; {��r�ꤤ�k��Y�z�Ӧr��,�h�ɦW�ᤧ�����ɦW�ɫܦn��}
var R:string;				       {�� _CUTRIGHT('TEST.BMP',4) �Y�Ǧ^'TEST'}
begin
  R:='';

  if (S<>'') and (L<=Length(S)) and (L>0) then
     begin
      R:=COPY(S,1,(Length(S)-L));
     end;

  if L=0 then R:=S;

  Result:=R;
end;
//==============================================================================
procedure _StrSwap(var S1,S2:string);
var S3:string;
begin
  S3:=S2;
  S2:=S1;
  S1:=S3;
end;
//==============================================================================
function _StrRange(S,R:string):Boolean; {�P�_ S �O�_�b R �϶��r�� �Y 'R1-R2' ��}
var N:integer;
    R1,R2,R3:string;
begin
    N:=_SubStrNum(R,',');
    if N>0 then {�p��d��r�ꤤ���X�ӳr�I}
       begin
       showmessage('�ǤJ _StrRANGE ��Ƥ��d��r��t���r�I');
       Result:=False;
       exit;
       end;

    N:=_SubStrNum(R,'-');
    if N>1 then {�p��d��r�ꤤ���X�Ӿ�b}
       begin
       showmessage('�ǤJ _StrRANGE ��Ƥ��d��r���b�ƿ��~');
       Result:=False;
       exit;
       end;
    if N=0 then
       begin
       if S=R then Result:=True
	 else
	  Result:=False;
       exit;
       end;

    R1:=_Left(R,(POS('-',R)-1));
    R2:=_CUTLeft(R,POS('-',R));

    if R1>R2 then
       begin
       R3:=R2;
       R2:=R1;
       R1:=R3;
       end;

    if (S>=R1) and (S<=R2) then
       Result:=True
       else
       Result:=False;

end;
//==============================================================================
function _StrMutiRANGE(S,R:string):Boolean;
var N,I:integer;
    SL:TstringLIST;
    RR:Boolean;
    R1,R2,SS:string;
begin
  SL:=TstringLIST.CREATE;
  RR:=False;
  _StrToStrlist(R,SL);


  for I := 0 to (SL.Count-1) do
  begin
    SS:=SL.stringS[I];
    N:=_SubStrNum(SS,'-');
    if N>1 then {�p��d��r�ꤤ���X�Ӿ�b}
       begin
       showmessage('�ǤJ _StrMuitRANGE ��Ƥ��d��r���b�ƿ��~');
       SL.FREE;
       Result:=False;
       exit;
       end;
  end;


  for I := 0 to (SL.Count-1) do
  begin
    SS:=SL.stringS[I];
    N:=_SubStrNum(SS,'-');

    if N=0 then     {�P_StrBelong()��ƥ\��}
       begin
       if S=SS then
	  begin
	    RR:=True;
	    break;
	  end;
       end
    else
       begin
	 R1:=_Left(SS,(POS('-',SS)-1));
	 R2:=_CUTLeft(SS,POS('-',SS));

	 if R1>R2 then _STRSWAP(R1,R2);

	 if (S>=R1) and (S<=R2) then
	    begin
	      RR:=True;
	      break;
	    end;
       end;

  end;

SL.free;
Result:=RR;

end;
//==============================================================================
function _StrBelong(S,R:string):Boolean; {���\���_StrMutiRANGE()��ƧP�_��i}
var I:integer;
    SL:TstringLIST;
    RR:Boolean;
    SS:string;
begin
  SL:=TstringLIST.CREATE;
  RR:=False;
  _StrToStrlist(R,SL);

  for I := 0 to (SL.Count-1) do
  begin
    SS:=SL.stringS[I];
    if _AllTrim(S)=_AllTrim(SS) then
       begin
	 RR:=True;
	 BREAK;
       end;
  end;

SL.FREE;
Result:=RR;

end;
//==============================================================================
function _StrLike(S,R:string):Boolean; {���\����_StrBelong()���,��������j�p�g}
var I:integer;
    SL:TstringLIST;
    RR:Boolean;
    SS:string;
begin
  SL:=TstringLIST.Create;
  RR:=False;
  _StrToStrList(R,SL);

  for I := 0 to (SL.Count-1) do
  begin
    SS:=SL.strings[I];
    if lowercase(_AllTrim(S))=lowercase(_AllTrim(SS)) then
       begin
	 RR:=True;
	 Break;
       end;
  end;

SL.Free;
Result:=RR;

end;
//==============================================================================
function _CutCharRight(S,C:string;F:integer):string; {��r��S��,�̥k��C�r���k��H�ᤧ�r�� ; F=0 �O�dC , F<>0 �sC �h��}
var R:string;
begin
  R:=S;
  if Length(C)<>1 then
     begin
     showmessage('_CutCharRight �ǤJ�Ѽƪ��׿��~');
     Result:=R;
     exit;
     end;

  if (S<>'') and (Pos(C,S)>0) then
     begin
      repeat
       if _Right(R,1)=C then Break;
       R:=_CutRight(R,1);
      until Length(R)=0;
     end;

  if (R<>'') and (F<>0) and (_Right(R,1)=C) then
      R:=_CutRight(R,1);

  Result:=R;

end;
//==============================================================================
function _CutAllChar(S:string;C:char):string;  {�N�r�� S ���Y��@�r�� C ���Ʊ���}
var I:integer;
    SS,R:string;
begin
  R:='';
  if Length(S)=0 then
     begin
       Result:='';
       exit;
     end;

  if Length(S)=1 then
     begin
       if S[1]=C then Result:=''
		 else Result:=S;
       exit;
     end;

  for I:=1 to Length(S) do
      begin
	SS:=Copy(S,I,1);
	if SS[1]=C then Continue
	   else R:=R+SS;
      end;

  Result:=R;

end;
//==============================================================================
function _SegNum(S,SEG:string):integer; {�P�_S�r�ꤤ���h��"�`" ( �Y SEG='' �h�H�r�I���` ) }
begin
  if S='' then
     begin
       Result:=0;
       exit;
     end;

  if SEG='' then SEG:=',';

  Result:=_SubStrNum(S,SEG)+1; {�`�� = �`�r���+1 }

end;
//==============================================================================
function _SubStrNum(S,SUBS:string):integer; {���r�ꤤ�X�{�Y�@�l�r�ꤧ���� }
var L,N,I : integer;
begin

  if (Length(S)=0) or (Length(SUBS)=0) then
      begin
	Result:=0;
	exit;
      end;

  N:=0;
  L:=Length(SUBS);
  {
  for I:=1 to (Length(S)-L+1) do
      begin
      if copy(S,I,L)=SUBS then N:=N+1;
      end;}

  i:=1;  {1999/06/21 �ץ�,�p���~���|�N_SubStrNum('EEE','EE') ��2,����1�~��(��L������) }
  while i<=(Length(S)-L+1) do
    begin
      if copy(S,I,L)=SUBS then
         begin
           N:=N+1;
           i:=i+L;
         end
      else
        i:=i+1;
    end;

  Result:=N;

end;
//==============================================================================
function _SegStr(S:string;I:integer):string; {���X S �r�ꤤ�� I �`�r��(�H�r�I��)}
var T :TstringLIST;
begin
  if (S='') or (I<=0) then
     begin
       {Result:='NIL';}
       Result:='';
       exit;
     end;

  T:=TStringLIST.Create;
  _StrToStrList(S,T);

  if I>T.Count then Result:=''
  else
    Result:=T[I-1];

  T.Free;
end;
{==============================================================================}
{���o S �r�ꤤ�H Seg �l�r����j���� N �Ӧr��Ϭq}
function _GetSegStr(S:string ; Seg:string ; N:integer):string; { ���q��!! }
var ss,tt:string;
    i:integer;
begin
  if (S='') or (Seg='') or (n=0) or (S=Seg) then
     begin
       Result:='';
       exit;
     end;

  if (pos(Seg,S)=0) and (n<>1) then
     begin
       Result:='';
       exit;
     end;

  if (pos(Seg,S)=0) and (n=1) then
     begin
       Result:=S;
       exit;
     end;

  ss:=S;
  for i:=1 to n do
      begin
        //showmessage(ss);

        if ss='' then
           begin
            tt:='';
            break;
           end;

        if pos(Seg,ss)<>0 then
           begin
           tt:=_Left(ss,(pos(Seg,ss)-1));
           end
        else
           begin
             if i=n then
                tt:=ss
             else
                tt:='';

             break;
           end;

        ss:=_cutLeft(ss,Length(tt)+Length(Seg));

      end;

  Result:=tt;

end;
//==============================================================================
function _GetCharSegStr(S:string ; C:Char ; N:integer):string; {���o S �r�ꤤ�H C �r���Ÿ����j���� N �Ӧr��Ϭq}
var T :TstringLIST;
    P,I : integer;
begin
  if (S='') or (N<=0) then
     begin
       Result:='';    {�Ѽƿ��~�@�߶Ǧ^�Ŧr��}
       exit;
     end;

  P:=_SubStrNum(S,C);    {�p��ǤJ�r�ꤤ���X���ѧO�Ÿ�}
  if P=0 then
     begin
       if N=1 then
	  begin
	    Result:=S;
	    exit;
	  end
       else
	  begin
	    Result:='';    {�ѼƤ��L�ѧO�Ÿ��BN>1�ɶǦ^�Ŧr��}
	    exit;
	  end;
     end;

  T:=TStringList.Create;

  for I := 1 to (P+1) do
    begin
      if I=(P+1) then
	 T.Add(S)
      else
	 T.Add(_Left(S,Pos(C,S)-1));

      S:=_CutLeft(S,Pos(C,S));
    end;

  if N>T.Count then Result:='' else
     Result:=T[N-1];

  T.Free;

end;
{==============================================================================}
procedure _ReplaceStr(var SourceStr:string ; OldStr,NewStr:string);
var i,skip:longint;
    r:string;
begin

  if Length(SourceStr)<Length(OldStr) then
     begin
       //Result:=SourceStr;
       exit;
     end;

  r:='';
  skip:=0;
  for i:=1 to (Length(SourceStr)-Length(OldStr)+1) do
      begin
        //_showint(i);
        if skip<>0 then
           begin
             skip:=skip-1;
             continue;
           end;
        if uppercase(copy(SourceStr,i,Length(OldStr)))=uppercase(OldStr) then
           begin
           r:=r+NewStr;
           skip:=Length(OldStr)-1;
           end
        else
           r:=r+copy(SourceStr,i,1);
      end;

  r:=r+_right(SourceStr,Length(OldStr)-1);

  SourceStr:=r;

end;
{==============================================================================}
procedure _ReplacePuncStr(var SourceStr:string); {�N���I�Ÿ��������j�g}
var i:longint;
    t,r:string;
begin

  if Length(SourceStr)=0 then
     begin
       //Result:=SourceStr;
       exit;
     end;

  r:='';
  for i:=1 to Length(SourceStr) do
      begin

        case copy(SourceStr,i,1)[1] of
          ',': t:='�A';
          ':': t:='�G';
          ';': t:='�F';
          '?': t:='�H';
          '!': t:='�I';
          '"': t:=' " ';
          '(': t:='�]';
          ')': t:='�^';
        else
          t:=copy(SourceStr,i,1)[1];
        end;

        r:=r+t;

      end;

  SourceStr:=r;

end;
{==============================================================================}
procedure _ReplaceSegStr(var S:string ; Seg:string ; SegIndex:integer ; ReplaceStr:string);
var i,segnum:integer;
    r:string;
begin
  segnum:=_SubStrNum(S,Seg)+1;
  if SegIndex>segnum then
     begin
       r:=S;
       for i:=1 to SegIndex-segnum do
           r:=r+Seg;
       r:=r+ReplaceStr;
     end;

  if SegIndex=segnum then
     begin
       r:='';
       for i:=1 to SegIndex-1 do
           r:=r+_GetSegStr(S,Seg,i)+Seg;
       r:=r+ReplaceStr;
     end;

  if SegIndex<segnum then
     begin
       r:='';
       for i:=1 to SegIndex-1 do
           r:=r+_GetSegStr(S,Seg,i)+Seg;

       r:=r+ReplaceStr+Seg;

       for i:=SegIndex+1 to segnum do
           r:=r+_GetSegStr(S,Seg,i)+Seg;
       r:=_cutright(r,Length(Seg));
     end;

  S:=r;

end;
{==============================================================================}
procedure _SegStrValueInc(var S:string ; Seg:string ; SegIndex:integer);
var v:LongInt;
begin
  v:=_StrToInt(_GetSegStr(S,Seg,SegIndex));
  v:=v+1;
  _ReplaceSegStr(S,Seg,SegIndex,inttostr(v));
end;
{==============================================================================}
function _FilterNumStr(S: string): string; { �N�r�ꤤ'�Ʀr�r��'�o�X }
var
   I : integer;
begin
   Result := '';

   for I := 1 to Length(S) do	  { �� '1' - '9' �� '-' �o�X }
       if (S[I] IN ['0'..'9']) or (S[I]='-') then
	  Result := Result + S[I];

   {�h���Ҧ���'-',�u�d�U�@�Ө��\�̫ܳe}
   {if _SubStrNum(Result,'-')<>0 then
      begin
      Result:=_CUTALLCHAR(Result,'-');
      Result:='-'+Result;
      end;}

    {87.06.09 �ק�}
    if _Left(Result,1)='-' then
       begin
         Result:=_cutallchar(Result,'-');
         Result:='-'+Result;
       end
    else Result:=_cutallchar(Result,'-');

end;
{==============================================================================}
function _StrRightSide(S:string ; L:integer):string; {�r��k�a}
var ss:string;
begin
if L<=0 then
   begin
     Result:='';
     exit;
   end;

ss:=_AllTrim(S);

if Length(ss)>=L then
   Result:=_right(ss,L)
else
   Result:=_space(L-Length(ss))+ss;

end;
{==============================================================================}
function _StrPos(S,subS:string):integer; {�D �r��subs �X�{�� �r��S ���ĴX��}
var r,i:integer;
begin
  r:=0;
  if Length(subs)>Length(S) then
     begin
       Result:=r;
       exit;
     end;

  if (S='') or (subs='') then
     begin
       Result:=r;
       exit;
     end;

  for i:=1 to (Length(S)-Length(subs)+1) do
      begin
        if copy(S,i,Length(subs))=subs then
           begin
             r:=i;
             break;
           end;
      end;

  Result:=r;

end;
{==============================================================================}
{�� �r��S ���X�{ �r��KeyWord ��s���Len�Ӧr��X��}
function _StrKeyWordValue(S,KeyWord:string ; Len:integer):string;
var r : string;
    i,l : integer;
begin
r:='';
i:=_StrPos(S,KeyWord);

if i=0 then
   begin
     Result:=r;
     exit;
   end;

l:=Len;
if (i+Length(KeyWord)+l-1)>Length(S) then
   l:=(Length(S)-(i+Length(KeyWord)-1));

{showmessage(inttostr(l)); }

Result:=copy(S,(i+Length(KeyWord)),l);

end;
//==============================================================================
// ***** TStrins ����
//==============================================================================
procedure _GetLeftStrings(S:Tstrings;L:integer; D:Tstrings);
var I : integer;
begin
D.Clear;

if L>S.Count then L:=S.Count;

for I := 0 to (L-1) do
    D.Add(S[I]);

end;
{==============================================================================}
procedure _GetRightStrings(S:Tstrings;L:integer; D:Tstrings);
var {T : Tstrings;}
    I : integer;
begin
{T:=TStringList.Create;
T.Clear; }
D.Clear;

if L>S.Count then L:=S.Count;

for I := (S.Count-L) to (S.Count-1) do
    D.Add(S[I]);

{
D.Clear;
D.Addstrings(T);

T.Free; }

end;
{==============================================================================}
procedure _DeleteStrings(S:Tstrings;P:integer);
var T : TstringList;
    I : integer;
begin
T:=TstringList.Create;
T.Clear;

if P>(S.Count) then
   begin
     T.Free;
     exit;
   end;

for I := 0 to (S.Count-1) do
    begin
    if I<>(P-1) then T.ADD(S[I]);
    end;

S.Clear;
S.AddStrings(T);

T.Free;

end;
{==============================================================================}
procedure _InsertStrings(S:Tstrings;P:integer;Str:string);
var T : TstringList;
    I : integer;
begin
T:=TstringList.Create;
T.Clear;

if P>(S.Count) then
   begin
     showmessage('_Insertstrings( ) ��ƴ��J��m���~�I');
     T.Free;
     exit;
   end;

for I := 0 to (S.Count-1) do
    begin
    if I=(P-1) then T.Add(Str);
    T.Add(S[I]);
    end;

S.Clear;
S.Addstrings(T);

T.Free;

end;
{==============================================================================}
function _AddStrings(S1,S2 : Tstrings):Tstrings;
begin
  S1.AddStrings(S2);
  Result:=S1;
end;
{==============================================================================}
procedure _TStringsSaveToINISec(TS:Tstrings ; FileName,SEC:string);
var i:integer;
begin
  if FileName='' then FileName:=_AppIni;
  _DelINISection(FileName,SEC);
  _WriteINIStr(FileName,SEC,'Count',inttostr(TS.Count));
  for i:=0 to TS.Count-1 do
      begin
        _WriteINIStr(FileName,SEC,'Line'+inttostr(i),TS.strings[i]);
      end;
end;
{==============================================================================}
procedure _TstringsLoadFromINISec(TS:Tstrings ; FileName,SEC:string);
var i,c:integer;
begin
  if FileName='' then FileName:=_AppIni;
  TS.clear;
  c:=_strtoint(_GetINIStr(FileName,SEC,'Count'));
  for i:=0 to c-1 do
      begin
        TS.add(_GetINIStr(FileName,SEC,'Line'+inttostr(i)));
      end;
end;
{==============================================================================}
function _SeekItemIndex(S:string;ITM:Tstrings):integer;
	 {���o�Y�@�r��S �b�@ Tstrings (�pComBoBox��ITEMS) ���̱��� INDEX ��}
var R,I:integer;
begin
  R:=-1;
  if Length(S)<>0 then
     begin
     for I:=0 to (ITM.Count-1) do
	 begin
	 if S=_Left(ITM.Strings[I],Length(S)) then
	    begin
	    R:=I;
	    Break;
	    end;
	 end;
     end;

  Result:=R;
end;
{==============================================================================}
function _SeekItemText(S:string;ITM:Tstrings):string;
	 {���o�Y�@�r��S �b�@ Tstrings (�pComBoBox��ITEMS) ���̱��񤧦r��}
var R:string;
    I:integer;
begin
  R:='';
  if Length(S)<>0 then
     begin
     for I:=0 to (ITM.Count-1) do
	 begin
	 if S=_Left(ITM.stringS[I],Length(S)) then
	    begin
	    R:=ITM.stringS[I];
	    BREAK;
	    end;
	 end;
     end;
  Result:=R;
end;
{==============================================================================}
function _ItemExist(T:TstringList;item:string):Boolean;
var i:integer;
    r:Boolean;
begin
  r:=False;
  for i:=0 to (T.Count-1) do
    begin
      if _AllTrim(uppercase(item))=_AllTrim(uppercase(T.strings[i])) then
         begin
           r:=True;
           break;
         end;
    end;
  Result:=r;
end;
{==============================================================================}
function _stringListValue(T:TstringList ; Para:string):string;
var r:string;
begin
  r:='';
  r:=T.Values[PARA];
  Result:=r;
end;
//==============================================================================
// ***** StringGrid ����䴩���
//==============================================================================
procedure _StrGridDelete(S: TstringGrid);
var I:longint;
begin

with TstringGrid(S) do
     begin
     if (rowCount-fixedrows)>1 then
	begin
	for I:=row to (rowCount-1) do
	    begin
	      rows[I]:=rows[I+1];
	    end;
	rowCount:=rowCount-1;
	end else
	    if (rowCount-fixedrows)=1 then  {�ܤ֫O�d�@��rows}
	       begin
		 rows[row].clear;
	       end;
     end;
end;
{==============================================================================}
procedure _StrGridInsert(S: TstringGrid);
var I:longint;
begin

with TstringGrid(S) do
     begin
       rowCount:=rowCount+1;
       for I:=(rowCount-1) downto (row+1) do
	   begin
	     rows[I]:=rows[I-1];
	   end;
       rows[row].clear;
     end;
end;
{==============================================================================}
procedure _StrGridAppend(S: TstringGrid);
begin

with TstringGrid(S) do
     begin
       rowCount:=rowCount+1;
       row:=(rowCount-1);
       rows[row].clear;
     end;
end;
{==============================================================================}
procedure _StrGridAdd(S: TstringGrid);
var I:longint;
begin

with TstringGrid(S) do
     begin
       rowCount:=rowCount+1;
       for I:=(rowCount-1) downto (row+2) do
	   begin
	     rows[I]:=rows[I-1];
	   end;
       rows[row+1].clear;
     end;
end;
{==============================================================================}
procedure _StrGridDelEmptyRow(S: TstringGrid);
var i:integer;
begin

if _IsInitStrGrid(S) then exit ; {�w�ūh�����P�_}

i:=S.Fixedrows;

while i<=(S.RowCount-1) do
      begin
	while (_IsEmptyRow(S,i)) do
	      begin
		{if i=(S.RowCount-1) then break;}
		_StrGridDeleteRow(S,i);
		{i:=s.row;  }
	      end;
	i:=i+1;
      end;

if S.RowCount=S.Fixedrows then
   _StrGridInit(S);
end;
{==============================================================================}
procedure _StrGridInit(S: TstringGrid);   {��l�� stringGrid �u�d�@ ROW }
begin
with TstringGrid(S) do
     begin
       rowCount:=Fixedrows+1;

       ROW:=FixedRows;
       COL:=FixedCols; { CELL ���ܥ��W }

       _StrGridClearRow(TstringGrid(S),(rowCount-1));
     end;

end;
{==============================================================================}
procedure _StrGridClear(S: TstringGrid);
var I:longint;
    O : TstringList;
begin
O := TstringList.Create;
O.Clear;

with TstringGrid(S) do
     begin
       for I:=fixedRows to (rowCount-1) do
	   begin
	     O.Clear;
	     _GetLeftstrings(rows[I],fixedCols,O);
	     rows[I].clear;
	     rows[I]:=O;
	   end;
     end;

O.Free;

end;
{==============================================================================}
procedure _StrGridDispLabel(S: TstringGrid ; DP: string);
var I : integer;
begin
with TstringGrid(S) do
     begin
       for I := 1 to colCount do
	   begin
	     cells[(I-1),0]:=_SegStr(DP,I);
	   end;
     end;
end;
{==============================================================================}
procedure _StrGridSetText(S: TstringGrid ; T:string);
begin
with TstringGrid(S) do
     begin
       CELLS[COL,ROW]:=T;
     end;
end;
{==============================================================================}
function _StrGridNextCell(S: TstringGrid):Boolean;
begin
Result:=True;

with TstringGrid(S) do
     begin
       if col<(colCount-1) then
	  begin
	  col:=col+1;
	  exit;
	  end;

       if col=(colCount-1) then
	  begin
	    if row<(rowCount-1) then
	       begin
		 row:=row+1;
		 col:=FixedCols;
		 exit;
	       end else
	       Result:=False;
	  end;

     end;
end;
{==============================================================================}
procedure _StrGridDeleteRow(S: TstringGrid ; R : LongInt);
var I:longint;
begin

with TstringGrid(S) do
     begin
     if (R>(ROWCount-1)) or (R<=(FixedRows-1)) then
	exit;	  { �`��ƥ~ �� �T�w�� ������ }

     if (rowCount-fixedrows)>1 then
	begin
	  for I:=R to (rowCount-1) do
	     begin
		rows[I]:=rows[I+1];
	     end;
	  rowCount:=rowCount-1;
	end else
	    if (rowCount-fixedrows)=1 then  {�ܤ֫O�d�@��rows}
		begin
		  rows[fixedrows].clear;
		end;

     end;

end;
{==============================================================================}
procedure _StrGridInsertRow(S: TstringGrid ; R : LongInt);
var I:longint;
begin

with TstringGrid(S) do
     begin
       if (R>(ROWCount-1)) or (R<=(FixedRows-1)) then
	exit;	  { �`��ƥ~ �� �T�w�� �����[ }

       rowCount:=rowCount+1;
       for I:=(rowCount-1) downto (R+1) do
	   begin
	     rows[I]:=rows[I-1];
	   end;
       rows[R].clear;
     end;
end;
{==============================================================================}
procedure _StrGridClearRow(S: TstringGrid ; R : LongInt);
var RR:TstringList;
begin
RR:=TstringList.Create;
RR.Clear;

with TstringGrid(S) do
     begin
     if (R>(ROWCount-1)) or (R<=(FixedRows-1)) then
	exit;	  { �`��ƥ~ �� �T�w�� �����M }

     _GetLeftstrings(Rows[R],FixedCols,RR);
     Rows[R].Clear;
     Rows[R]:=RR;
     end;

RR.Free;

end;
{==============================================================================}
function  _IsEmptyRow(S: TstringGrid ; R : LongInt):Boolean;
var RR:string;
    GG:TstringList ;
begin
GG:=TstringList.Create;
GG.Clear;
Result:=False;

with TstringGrid(S) do
     begin
     if (R>(ROWCount-1)) or (R<=(FixedRows-1)) then
	begin
	GG.Free;
	exit;	  { �`��ƥ~ �� �T�w�� ������ }
	end;

     _GetRightstrings(Rows[R],(ColCount-FixedCols),GG);
     _StrlistToStr(GG,RR);
     if _AllTrim(_CutAllChar(rr,','))='' then Result:=True;

     end;

GG.Free;

end;
{==============================================================================}
function  _IsInitStrGrid(S: TstringGrid ):Boolean;
begin
Result:=False;
with S do
     begin
     if (RowCount=Fixedrows) then  Result:=True;
     if (RowCount=Fixedrows+1) and _IsEmptyRow(S,(RowCount-1))
	then  Result:=True;
     end;

end;
{==============================================================================}
function  _IsLastRow(S: TstringGrid ):Boolean;	{�O�_�b�̫�@�� ROW }
begin
Result:=False;
with TstringGrid(S) do
     begin
     if row=(rowCount-1) then  Result:=True;
     end;

end;
{==============================================================================}
{�n���S�O�T�w��(fixcol),�C(fixrow)���ƥ�}
procedure _StrGridSaveToFile(StrG: TstringGrid ; FileName: string);
var
  i, j: LongInt;
  ss: string;
  f: TextFile;
begin
  AssignFile(f, FileName);
  Rewrite(f);
  ss := inttostr(StrG.ColCount) + ',' + inttostr(StrG.RowCount);
  Writeln(f, ss);

  for i := 0 to (StrG.RowCount - 1) do
  begin
    for j := 0 to (StrG.ColCount - 1) do
    begin
      if _RTrim(StrG.Cells[j, i]) <> '' then
      begin
	ss := inttostr(j) + ',' + inttostr(i) + ',' + StrG.Cells[j, i];
	Writeln(f, ss);
      end;
    end;
  end;
  CloseFile(f);
end;
{==============================================================================}
procedure _StrGridLoadFromFile(StrG: TstringGrid ; FileName: string);
var ss: string;
    f: TextFile;
begin
  if not FileExists(FileName) then exit;

  AssignFile(f, FileName);
  Reset(f);
  {if IOResult <> 0 then
     begin
       showmessage('File ' + FileName + ' not found');
       exit;
     end;
   }
  Readln(f, ss);
  if ss <> '' then
  begin
    StrG.ColCount := _strtoint(_SegStr(ss,1));
    StrG.RowCount := _strtoint(_SegStr(ss,2));
    _StrGridClear(StrG);
  end;

  while not Eof(f) do
  begin
    Readln(f, ss);
    StrG.Cells[_strtoint(_SegStr(ss,1)),
	       _strtoint(_SegStr(ss,2))] :=_SegStr(ss,3);
  end;
  CloseFile(f);
end;
{==============================================================================}
procedure _StrGridSaveToLines(StrG:TstringGrid ; L:Tstrings ; formatStr:string);
var I,J:integer;
    S:string;
begin				  {�T�w�C���s}
  L.clear;

  if _IsInitStrGrid(StrG) then exit ; {�w�ūh�����P�_}

  for I := StrG.FixedRows to (StrG.RowCount-1) do
      begin
	S:='';
	for J := 1 to _SubStrNum(FormatStr,',') do
	    begin
	      S:=S+_ReStr(StrG.Rows[I][(StrG.FixedCols-1)+J],
		   _strtoint(_SegStr(FormatStr,J)));
	    end;
	L.ADD(S);
      end;
end;
{==============================================================================}
procedure _StrGridLoadFromLines(StrG:TstringGrid ; L:Tstrings ; formatStr:string);
var I,J:integer;
    S:string;
    FixR : TstringList;
    R : TstringList;
begin
  if L.Count=0 then
     begin
     {_STRGRIDINIT(STRG); }
     exit;
     end;

  FixR:=TstringList.Create;
  R:=TstringList.Create;
  FixR.Clear;
  R.Clear;

  StrG.RowCount:=StrG.FixedRows + L.Count;
  _StrGridClear(StrG);

  for I := 0 to (L.Count-1) do
      begin
       FixR.Clear;
       R.Clear;
       S:=L[I];
       for J:= 1 to _SubStrNum(FormatStr,',') do
	    begin
	      R.ADD(_Left(S,_strtoint(_SegStr(FormatStr,J))));
	      S:=_CUTLeft(S,_strtoint(_SegStr(FormatStr,J)))
	    end;

	_GetLeftstrings(StrG.Rows[StrG.FixedRows+I],StrG.FixedCols,FixR);
	StrG.Rows[StrG.FixedRows+I]:=_Addstrings(FixR,R);
      end;

  FixR.Free;
  R.Free;

end;
{==============================================================================}
procedure _StrGridSaveToINISec(StrG: TstringGrid ; FileName,Sec: string ; all:integer);
var i, j ,ii,jj: LongInt;
    ss: string;
begin
  _DelINISection(FileName,Sec);

  _WriteINIStr(FileName,Sec,'ColCount',inttostr(StrG.ColCount));
  _WriteINIStr(FileName,Sec,'RowCount',inttostr(StrG.RowCount));

  ii:=0 ; jj:=0;
  if all=0 then
     begin
      ii:=StrG.FixedRows;
      JJ:=StrG.FixedCols;
      if (ii=StrG.RowCount) or (jj=StrG.ColCount) then
	 exit;
     end;


  for i := ii to (StrG.RowCount - 1) do
  begin
    for j := jj to (StrG.ColCount - 1) do
    begin
      if _RTRIM(StrG.Cells[j, i]) <> '' then
      begin
	{ss := inttostr(j) + ',' + inttostr(i) + ',' + StrG.Cells[j, i];
	Writeln(f, ss); }
	ss:=inttostr(j) + '_' + inttostr(i);
	_WriteINIStr(FileName,Sec,ss,StrG.Cells[j, i]);
      end;
    end;
  end;
end;
{==============================================================================}
procedure _StrGridLoadFromINISec(StrG: TstringGrid ; FileName,Sec: string);
var X,Y,CellX,CellY: integer;
    Cell,CellText: string;
    SecBuf:TstringList;
begin
  if _GetIniSecKeyNum(FileName,Sec)=0 then exit;

  SecBuf:=TstringList.Create;
  SecBuf.Clear;

  X:=_strtoint(_GetINIStr(FileName,Sec,'ColCount'));
  Y:=_strtoint(_GetINIStr(FileName,Sec,'RowCount'));
  if X>0 then StrG.ColCount:=X else
     begin
       showmessage('INI�ɤ�ColCount�ȿ��~�I');
       exit;
     end;
  if Y>0 then StrG.RowCount:=Y else
     begin
       showmessage('INI�ɤ�RowCount�ȿ��~�I');
       exit;
     end;

  _StrGridClear(StrG);
  _GetINISection(FileName,Sec,SecBuf);

  _Deletestrings(SecBuf,1);
  _Deletestrings(SecBuf,1); {�s����զr��(��C�Ƹ�T)}

  for X:= 1 to SecBuf.Count do
      begin
	Cell:=_GetCharSegStr(SecBuf[x-1],'=',1);
	CellText:=_GetCharSegStr(SecBuf[x-1],'=',2);
	CellX:=_strtoint(_GetCharSegStr(Cell,'_',1));
	CellY:=_strtoint(_GetCharSegStr(Cell,'_',2));

	StrG.Cells[CellX,CellY]:=CellText;
      end;

  {Memo1.Clear;
  DelphiIni.ReadSectionValues('Transfer', Memo1.Lines);
  Memo1.Lines.Values['Title1'] := 'Picture Painter';  }

end;
{==============================================================================}
{�u�O�C�@�C���e���col,�ΨӰO�ѼƳ]�w��}
{Example: _StrGridGetINISec(stringGrid1,'d:\ticket\ticket.ini','print');       }
procedure _StrGridGetINISec(StrG:TstringGrid ; INIFile,Section:string);
var j:integer;
begin
  StrG.fixedcols:=1;
  StrG.fixedrows:=0;
  StrG.rowCount:=_GetINISecKeyNum(INIFile,Section);
  StrG.colCount:=2;
  _GetINISectionKeys(INIFile,Section,StrG.Cols[0],0);
  _GetINISectionKeys(INIFile,Section,StrG.Cols[1],1);

  for j := 0 to (StrG.rowCount-1) do
    begin
      StrG.cells[1,j]:=_cutLeft(StrG.cells[1,j],pos('=',StrG.cells[1,j]));
    end;
end;
{==============================================================================}
{�u�O�C�@�C���e���col,�ΨӰO�ѼƳ]�w��}
{Example: _StrGridWriteINISec(stringGrid1,'d:\ticket\ticket.ini','print');     }
procedure _StrGridWriteINISec(StrG:TstringGrid ; INIFile,Section:string);
var j:integer;
    key,para:string;
begin
  for j := 0 to (StrG.rowCount-1) do
      begin
        key:=StrG.cells[0,j];
        para:=StrG.cells[1,j];
        if _AllTrim(key)<>'' then
           begin
             _WriteINIStr(INIFile,Section,key,para);
           end;
      end;
end;
{==============================================================================}
function _StrGridText(StrG:TstringGrid):string;
begin
  Result:=StrG.cells[StrG.col,StrG.row];
end;
{==============================================================================}
procedure _StrGridTextWrite(StrG:TstringGrid ; Text:string);
begin
  StrG.cells[StrG.col,StrG.row]:=Text;
end;
{==============================================================================}
function _StrGridCellRect(Sender: TObject; ACol, ARow: Longint):TRect;
var T:TRect;
begin
{  �� stringGrid �� SelectCell �ƥ󤤩I�s ( �Ǧ^�� CELL ����m�d�� )	  }
{  DBLookupcombo1.BoundsRect:=_StrGridCellRect(Sender, Col, Row);   }

  if not (SendER IS TstringGrid) then exit;

  With TstringGrid(sender) do
       begin
	 T.Left:=CellRect(ACol,ARow).Left+Left;
	 T.Top:=CellRect(ACol,ARow).Top+Top;
	 T.Right:=CellRect(ACol,ARow).Right+Left+2;
	 T.Bottom:=CellRect(ACol,ARow).Bottom+Top+2;
       end;

  Result:=T;

end;
{==============================================================================}
function _StrGridGotoKeyRow(SG:TstringGrid ; key_field:integer ; key:string):Boolean;
var j,rowindex : integer;
begin

  rowindex:=-1;
  for j:=0 to SG.rowCount-1 do
      begin
        //_SAVELOG(inttostr(KEY_FIELD-1)  ,0);
        if SG.cells[key_field-1,j]=key then rowindex:=j;
      end;

  // _SAVELOG(inttostr(rowindex),0);
  SG.row:=rowindex;
  //_SAVELOG(inttostr(SG.ROW),0);

  if rowindex=-1 then
     Result:=False
  else
     Result:=True;

end;
//==============================================================================
// ***** INI �� & counter �B�z
//==============================================================================
function _GetINIStr0(FileName,Section,KeyName:string):string; //�ª� ; for Delphi 2.0 , ����
var GETCHR:PCHAR;
    GETINT:integer;
    S,K,F:PCHAR;
begin
  GETCHR:=_StrTOPchar(_SPACE(255));

  S:=_StrTOPchar(SECTION);
  K:=_StrTOPchar(KEYNAME);
  F:=_StrTOPchar(FILENAME);

  GETINT:=GETPRIVATEPROFILEstring(S,K,
		  '',GETCHR,255,F);

  Result:=_AllTrim(STRPAS(GETCHR));

  STRDISPOSE(GETCHR); //��GETCHR='' ��,��remark �b���ǵ{���|���
  STRDISPOSE(S);
  STRDISPOSE(K);
  STRDISPOSE(F);
end;
//==============================================================================
function _WriteINIStr0(FileName,Section,KeyName,Parameter:string):Boolean; //�ª� for Delphi 2.0 , ����
var S,K,F,P:PCHAR;
begin
  {
  S:=_StrTOPchar(Section);
  K:=_StrTOPchar(KeyName);
  F:=_StrTOPchar(FileName);
  P:=_StrTOPchar(Parameter);

  Result:=WRITEPRIVATEPROFILEstring(S,K,P,F);

  STRDISPOSE(S);
  STRDISPOSE(K);
  STRDISPOSE(F);
  STRDISPOSE(P);
  }

  //Parameter �̦h�u��� 254 �r��
  Result:=WritePrivateProFileString(PChar(Section),PChar(KeyName),PChar(Parameter),PChar(FileName));

end;
//==============================================================================
function _GetINIStr(FileName,Section,KeyName:string):string;
var MyIni: TIniFile;
    R:string;
begin
  MyIni := TIniFile.Create(PChar(FileName));
  R:=MyIni.ReadString(PChar(Section),PChar(KeyName), '');
  MyIni.Free;
  Result:=R;
end;
//==============================================================================
function _WriteINIStr(FileName,Section,KeyName,Parameter:string):Boolean;
var MyIni: TIniFile;
begin
  MyIni := TIniFile.Create(PChar(FileName));
  MyIni.WriteString(PChar(Section),PChar(KeyName),PChar(Parameter));
  MyIni.Free;

  Result:=True; //�û��Ǧ^ True , ���M�¨�ƼҦ�
end;
//==============================================================================
function _ReadWriteINIStr(FILE_NAME,SEC,KEYNAME,DefauStr:string):string;
var R:string;
begin
R:=_GetINIStr(FILE_NAME,SEC,KEYNAME);
if _RTRIM(R)<>'' then
   begin
     Result:=R;
   end else
   begin
     R:=DefauStr;
     _WriteINIStr(FILE_NAME,SEC,KEYNAME,R);
     Result:=R;
   end;

end;
//==============================================================================
procedure _DelINISection(FILE_NAME,SEC:string);
var
  MyIni: TIniFile;
begin
  MyIni := TIniFile.Create(FILE_NAME);
  MyIni.EraseSection(SEC);
  MyIni.Free;
end;
//==============================================================================
procedure _GetINISection(FILE_NAME,SEC:string; T:Tstrings);
var MyIni: TIniFile;
begin
  MyIni := TIniFile.Create(FILE_NAME);
  T.Clear;
  MyIni.ReadSectionValues(SEC, T);
  MyIni.Free;
end;
//==============================================================================
procedure _GetINISecKeyWord(FILE_NAME,SEC:string; T:Tstrings);
var MyIni: TIniFile;
begin
  MyIni := TIniFile.Create(FILE_NAME);
  T.Clear;
  MyIni.ReadSection(SEC, T);
  MyIni.Free;
end;
//==============================================================================
function _GetINISecKeyNum(FILE_NAME,SEC:string):integer;
var T:TstringList;
begin
T:=TstringList.Create;
T.Clear;
_GetINISecKeyWord(FILE_NAME,SEC,T);
Result:=T.Count;
T.Free;
end;
{==============================================================================}
function _LaunchCounter:integer;
var c:integer;
begin
  c:=_readint('','','LaunchCounter',0);
  c:=c+1;
  _WriteInt('','','LaunchCounter',c);
  Result:=c;
end;
{==============================================================================}
function _CGICounter(CounterName:string):integer;
var c:integer;
begin
  c:=_readint('','',CounterName,0);
  c:=c+1;
  _WriteInt('','',CounterName,c);
  Result:=c;
end;
{==============================================================================}
procedure _TimeHit(EveryDay:integer); {���G�ɶ��O���s�J INI �ɤ� , EveryDay<>0 �ݨC�魫�m}
var T,TimeHit,D:string;
    Hour:integer;
begin
  T:='';
  _ReplaceSegStr(T,',',25,_date(0,'/')); {Default}
  TimeHit:=_QRStr('TimeHit',T);

  if EveryDay<>0 then {�ݨC�魫�m}
     begin
       D:=_GetSegStr(TimeHit,',',25);
       if D<>_date(0,'/') then
          begin
            _QWStr('TimeHit',T);
            TimeHit:=T;
          end;
     end;

  Hour:=_StrToInt(_WhatTime('Hour'));
  _SegStrValueInc(TimeHit,',',Hour+1); {00:??:?? �p�ɦb�Ĥ@�`}
  _QWStr('TimeHit',TimeHit);

end;
{==============================================================================}
procedure _GetINISectionNames(INIFile:string ; T:Tstrings);
var FileHandle:textfile;
    s:string;
begin
  T.clear;
  if not _file(INIFile) then exit;
  AssignFile(FileHandle,INIFile);
  ReSet(FileHandle);

  while not eof(FileHandle) do
    begin
      readln(FileHandle,s);
      if copy(s,1,1)='[' then
         T.add(s+' �`��') ;
    end;

  {
  repeat
  readln(FileHandle,s);
  if copy(s,1,1)='[' then
  T.add(s+' �`��');
  until seekeof(FileHandle);  }

  CloseFile(FileHandle);
end;
{==============================================================================}
{Example: _GetINISectionKeys('d:\ticket\ticket.ini','print',ListBox1.items,1); }
procedure _GetINISectionKeys(INIFile,Section:string ; T:Tstrings ; Val:integer);
var MyIni: TIniFile;
begin
  MyIni := TIniFile.Create(INIFile);
  T.Clear;

  if Val=0 then
     MyIni.ReadSection(Section,T)        {���tKEY ����}
  else
     MyIni.ReadSectionValues(Section,T); {�tKEY ���� (�����k�䤧��) }

  MyIni.Free;
end;
{==============================================================================}
{Example: _SectionCopy(_AliasAppIni('TICKET_R'),_appini,'Variable',1));        }
{                         Update=0 ���R�¸`�ϦA�[�J , Update<>0 ������s�¸`�� }
procedure _SectionCopy(SourceINI,DestINI,Section:string ; Update:integer);
var T:TstringList;
    i:integer;
    key,para:string;
begin
  if not _file(SourceINI) then
     begin
       showmessage('INI �ɤ��s�b -->' + SourceINI);
       exit;
     end;

  if (Update=0) and (_file(DestINI)) then _DelINISection(DestINI,Section);

  T:=TstringList.Create;
  _GetINISectionKeys(SourceINI,Section,T,1);

  for i :=0 to (T.Count-1) do
      begin
        key:=_cutcharright(T.strings[i],'=',1);
        para:=T.Values[key];
        _WriteINIStr(DestINI,Section,key,para);
      end;

  T.free;
end;
{==============================================================================}
function _ReadPassWord(IniFile,Section,VarName,DefauVal:string):string; {INI KeyWord �r��ѽX}
var P,R:string;

begin
   P:=_ReadStr(IniFile,Section,VarName,DefauVal);

   if _left(P,2)='&&' then
      begin
        P:=_CutLeft(P,2);
        R:=_encrypt(P);
      end
   else
      begin
        R:=P;
        P:=_encrypt(P);
        if P<>'' then _WriteStr(IniFile,Section,VarName,'&&'+P) ;
               //  else _WriteStr(IniFile,Section,VarName,'');
      end;

   Result:=R;

end;
{==============================================================================}
function _ReadStr(IniFile,Section,VarName,DefauVal:string):string;
begin
  if IniFile='' then IniFile:=_AppIni;
  if Section='' then Section:='Parameter';
  Result:=_ReadWriteINIStr(IniFile,Section,VarName,DefauVal);
end;
{==============================================================================}
function _WriteStr(IniFile,Section,VarName,VarValue:string):Boolean;
begin
  if IniFile='' then IniFile:=_AppIni;
  if Section='' then Section:='Parameter';
  Result:=_WriteINIStr(IniFile,Section,VarName,VarValue);
end;
{==============================================================================}
function _ReadInt(IniFile,Section,VarName:string ; DefauVal:integer):integer;
begin
  if IniFile='' then IniFile:=_AppIni;
  if Section='' then Section:='Parameter';
  Result:=_strtoint(_ReadWriteINIStr(IniFile,Section,VarName,inttostr(DefauVal)));
end;
{==============================================================================}
function _WriteInt(IniFile,Section,VarName:string ; VarValue:integer):Boolean;
begin
  if IniFile='' then IniFile:=_AppIni;
  if Section='' then Section:='Parameter';
  Result:=_WriteINIStr(IniFile,Section,VarName,inttostr(VarValue));
end;
{==============================================================================}
function _ReadBoolean(IniFile,Section,VarName:string ; DefauVal:Boolean):Boolean;
var d,r:string;
begin
  if DefauVal=True then d:='True'
     else d:='False';

  if IniFile='' then IniFile:=_AppIni;
  if Section='' then Section:='Parameter';

  r:=uppercase(_ReadWriteINIStr(IniFile,Section,VarName,d));

  if r='True' then
     begin
       _WriteINIStr(IniFile,Section,VarName,'True');
       Result:=True;
     end
  else
     begin
       _WriteINIStr(IniFile,Section,VarName,'False');
       Result:=False;
     end;

end;
{==============================================================================}
function _WriteBoolean(IniFile,Section,VarName:string ; VarValue:Boolean):Boolean;
begin
  if IniFile='' then IniFile:=_AppIni;
  if Section='' then Section:='Parameter';

  if VarValue=True then
     Result:=_WriteINIStr(IniFile,Section,VarName,'True')
  else
     Result:=_WriteINIStr(IniFile,Section,VarName,'False');

end;
{==============================================================================}
function _QRStr(VarName,DefauVal:string):string; {Quick ReadStr}
begin
  Result:=_ReadStr('','',VarName,DefauVal);
end;
{==============================================================================}
function _QWStr(VarName,VarValue:string):Boolean; {Quick WriteStr}
begin
  Result:=_WriteStr('','',VarName,VarValue);
end;
{==============================================================================}
function _QRInt(VarName:string ; DefauVal:integer):integer; {Quick ReadInt}
begin
  Result:=_ReadInt('','',VarName,DefauVal);
end;
{==============================================================================}
function _QWInt(VarName:string ; VarValue:integer):Boolean; {Quick WriteInt}
begin
  Result:=_WriteInt('','',VarName,VarValue);
end;
//==============================================================================
// ***** ��r�ɳB�z
//==============================================================================
function _LoadFromFile(FileName:string;P:integer):string;
var F : TextFile;
    S : string;
begin
  Result:='';
  if FileExists(FileName) then
     begin
       AssignFile(F, FileName);
       Reset(F);

       if P=0 then
          Readln(F, S)
       else
          while not Eof(f) do Readln(F, S);

       Result:=S;
       CloseFile(F);
     end;
end;
//==============================================================================
procedure _SaveToFile(FileName,S:string);  {�л\}
var F : TextFile;
begin
  _NewFile(FileName);
  AssignFile(F, FileName);
  Rewrite(F);
  Writeln(F, S);
  CloseFile(F);
end;
//==============================================================================
procedure _AppendToFile(FileName,S:string);
var F : TextFile;
begin
  if not _File(FileName) then _NewFile(FileName);
  AssignFile(F, FileName);
  Append(F);
  Writeln(F, S);
  CloseFile(F);
end;
//==============================================================================
procedure _NewFile(FileName:string);
var FileHandle : integer;
begin
  if FileExists(pchar(FileName)) then
     DeleteFile(pchar(FileName));
  FileHandle:=FileCreate(pchar(FileName));
  FileClose(FileHandle);
end;
//==============================================================================
procedure _DelFile(FileName:string);
begin
  if FileExists(pchar(FileName)) then
     DeleteFile(pchar(FileName));
end;
//==============================================================================
function _File(FileName:string):Boolean;
begin
  Result:=FileExists(pchar(FileName));
end;
//==============================================================================
function _FileLineCount(FileName:string):LongInt; {�D�ɮפ���Ʀ��}
var r:LongInt;
    FileHandle : Textfile ;
    tmp:string;
begin
  r:=0;
  if FileExists(FileName) then
     begin
       AssignFile(FileHandle, FileName);
       Reset(FileHandle);
       while not eof(FileHandle) do
         begin
           readln(FileHandle,tmp);
           r:=r+1;
         end;
       CloseFile(FileHandle);
     end;

  Result:=r;

end;
//==============================================================================
// ***** ��� DUMP , �ɮ���X
//==============================================================================
procedure _TableToExcel(TB:TDataSet);
var i,j,x,y:longint;
    msExcel:variant;
begin

  try
    msExcel:=createoleobject('excel.application');
  except
    showmessage('Don''t Start MS-Excel !');
    exit;
  end;

  msExcel.application.visible:=true;
  //msExcel.application.workbooks.open('c:\dump.xls');  {�}�Ҥw�s�b����}
  msExcel.application.workbooks.add;    {�}�s�ɮ�}

  j:=0;
  TB.first;
  while not TB.eof do
        begin
          for i := 0 to TB.FieldCount-1 do
              begin
                 //msExcel.worksheets('sheet1').cells[j+1,i+1].value:=TB.Fields[i].displaytext;  // Excel 95
                 msExcel.worksheets[1].cells[j+1,i+1].value:=TB.Fields[i].displaytext;    // Excel 97
              end;
          j:=j+1;
          TB.Next;
        end;
  

  //�����p��]�w�̾A�C���γ̾A��e

  {//�۰ʦs��
  try
    msExcel.activeworkbook.saveas('c:\dump.xls');
  except
    showmessage('Don''t SaveTo C:\dump.xls');
  end; }

  //msExcel.quit;  {����Excel�{��}

end;
{==============================================================================}
{Table �ɦL����r��,�i�]�w�k�a���(�ƭ���) }
procedure _TableDump(TB:TDataSet ; IntField:string);
var i,j,l:integer;
    FileHandle : textfile;
    FileName,tmp,dash : string;
    MyBookmark: TBookmark;
begin
  FileName:='C:\Dump.txt';
  if _AllTrim(FileName)='' then exit;

  if TB.active=False then
     begin
       showmessage('��ƪ��}�ҡA�L�k����ɦL�ʧ@�I');
       exit;
     end;

  assignfile(FileHandle,FileName);
  rewrite(FileHandle);


  {�Y}
  tmp:='';
  dash:='';
  j:=0;
  for i := 0 to (TB.fieldCount-1) do
      begin
        if TB.fields[i].visible=False then continue;
        j:=j+1;
        l:=TB.fields[i].displaywidth;
        if Length(TB.fields[i].displaylabel)>l then l:=Length(TB.fields[i].displaylabel);

        if _strlike(inttostr(j),IntField) then
             tmp:=tmp+_StrRightSide(TB.fields[i].displaylabel,l)+' '
        else tmp:=tmp+_restr(TB.fields[i].displaylabel,l)+' ';

        dash:=dash+_repeatstr('-',l)+' ';
      end;
  writeln(FileHandle,tmp);
  writeln(FileHandle,dash);


  {��}
  MyBookmark := TB.GetBookMark;
  TB.disablecontrols;
  TB.first;
  while not TB.eof do
     begin
       tmp:='';
       j:=0;
       for i := 0 to (TB.fieldCount-1) do
           begin
             if TB.fields[i].visible=False then continue;
             j:=j+1;
             l:=TB.fields[i].displaywidth;
             if Length(TB.fields[i].displaylabel)>l then l:=Length(TB.fields[i].displaylabel);

             if _strlike(inttostr(j),IntField) then
                  tmp:=tmp+_StrRightSide(TB.fields[i].displaytext,l)+' '
             else tmp:=tmp+_restr(TB.fields[i].displaytext,l)+' ';

           end;
       writeln(FileHandle,tmp);
       TB.next;
     end;

  TB.GotoBookMark(MyBookMark);
  TB.FreeBookmark(MyBookmark);
  TB.enablecontrols;
  CloseFile(FileHandle);

  {showmessage('�ɮ� '+FileName+' ���͡I');}
  _run('notepad.exe '+FileName);

end;
{==============================================================================}
{Table �ɦL����r��,�i������ɦL }
procedure _TableSelectDump(TB:TDataSet ; SelectField:string);
var i,j,l:integer;
    FileHandle : textfile;
    FileName,tmp,dash : string;
    MyBookmark: TBookmark;
begin
  FileName:='C:\Dump.txt';
  if _AllTrim(FileName)='' then exit;

  if TB.active=False then
     begin
       showmessage('��ƪ��}�ҡA�L�k����ɦL�ʧ@�I');
       exit;
     end;

  assignfile(FileHandle,FileName);
  rewrite(FileHandle);


  {�Y}
  tmp:='';
  dash:='';
  j:=0;
  for i := 0 to (TB.fieldCount-1) do
      begin
        if TB.fields[i].visible=False then continue;
        if not _strlike(inttostr(i+1),SelectField) then continue;
        j:=j+1;
        l:=TB.fields[i].displaywidth;
        if Length(TB.fields[i].displaylabel)>l then l:=Length(TB.fields[i].displaylabel);
        {
        if _strlike(inttostr(j),IntField) then
             tmp:=tmp+_StrRightSide(TB.fields[i].displaylabel,l)+' '
        else} tmp:=tmp+_restr(TB.fields[i].displaylabel,l)+' ';

        dash:=dash+_repeatstr('-',l)+' ';
      end;
  writeln(FileHandle,tmp);
  writeln(FileHandle,dash);


  {��}
  MyBookmark := TB.GetBookMark;
  TB.disablecontrols;
  TB.first;
  while not TB.eof do
     begin
       tmp:='';
       j:=0;
       for i := 0 to (TB.fieldCount-1) do
           begin
             if TB.fields[i].visible=False then continue;
             if not _strlike(inttostr(i+1),SelectField) then continue;
             j:=j+1;
             l:=TB.fields[i].displaywidth;
             if Length(TB.fields[i].displaylabel)>l then l:=Length(TB.fields[i].displaylabel);
             {
             if _strlike(inttostr(j),IntField) then
                  tmp:=tmp+_StrRightSide(TB.fields[i].displaytext,l)+' '
             else} tmp:=tmp+_restr(TB.fields[i].displaytext,l)+' ';

           end;
       writeln(FileHandle,tmp);
       TB.next;
     end;

  TB.GotoBookMark(MyBookMark);
  TB.FreeBookmark(MyBookmark);
  TB.enablecontrols;
  CloseFile(FileHandle);

  {showmessage('�ɮ� '+FileName+' ���͡I');}
  _run('notepad.exe '+FileName);

end;
{==============================================================================}
{Table �ɦL����r��(�i�]���`�p��������)}
procedure _TableCountDump(TB:TDataSet ; CountField:string);
var i,j,k,l:integer;
    FileHandle : textfile;
    FileName,tmp,dash,sumtmp : string;
    MyBookmark: TBookmark;
    {sum: array[1.._Max(_SubStrNum(CountField,','),1)] of Longint ;  }
    sum: array[1..20] of Longint ;
label quit;
begin
  FileName:='C:\Dump.txt';
  if _AllTrim(FileName)='' then exit;

  for i := 1 to 20 do sum[i]:=0;

  if TB.active=False then
     begin
       showmessage('��ƪ��}�ҡA�L�k����ɦL�ʧ@�I');
       exit;
     end;

  assignfile(FileHandle,FileName);
  rewrite(FileHandle);


  {�Y}
  tmp:='';
  dash:='';
  j:=0;
  for i := 0 to (TB.fieldCount-1) do
      begin
        if TB.fields[i].visible=False then continue;
        j:=j+1;
        l:=TB.fields[i].displaywidth;
        if Length(TB.fields[i].displaylabel)>l then l:=Length(TB.fields[i].displaylabel);

        if _strlike(inttostr(j),CountField) then
             tmp:=tmp+_StrRightSide(TB.fields[i].displaylabel,l)+' '
        else tmp:=tmp+_restr(TB.fields[i].displaylabel,l)+' ';

        dash:=dash+_repeatstr('-',l)+' ';
      end;
  writeln(FileHandle,tmp);
  writeln(FileHandle,dash);


  {��}
  MyBookmark := TB.GetBookMark;
  TB.disablecontrols;
  TB.first;
  while not TB.eof do
     begin
       tmp:='';
       j:=0;
       for i := 0 to (TB.fieldCount-1) do
           begin
             if TB.fields[i].visible=False then continue;
             j:=j+1;
             l:=TB.fields[i].displaywidth;
             if Length(TB.fields[i].displaylabel)>l then l:=Length(TB.fields[i].displaylabel);

             if _strlike(inttostr(j),CountField) then
                tmp:=tmp+_StrRightSide(TB.fields[i].displayText,l)+' '
             else tmp:=tmp+_restr(TB.fields[i].displayText,l)+' ';

             if _AllTrim(CountField)='' then continue;
             for k:=1 to _SubStrNum(CountField,',') do
                 begin
                   if _strtoint(_SegStr(CountField,k))=0 then continue;
                   if _strtoint(_SegStr(CountField,k))=j then
                      sum[k]:=sum[k]+_strtoint(TB.fields[i].displayText);
                 end;
           end;
       writeln(FileHandle,tmp);
       TB.next;
     end;

  if _AllTrim(CountField)='' then goto quit;

  {��}
  tmp:='';
  dash:='';
  j:=0;
  for i := 0 to (TB.fieldCount-1) do
      begin
        if TB.fields[i].visible=False then continue;
        j:=j+1;
        l:=TB.fields[i].displaywidth;
        if Length(TB.fields[i].displaylabel)>l then l:=Length(TB.fields[i].displaylabel);
        sumtmp:='';
        for k:=1 to _SubStrNum(CountField,',') do
                 begin
                   if _strtoint(_SegStr(CountField,k))=0 then continue;
                   if _strtoint(_SegStr(CountField,k))=j then
                      begin
                      sumtmp:=inttostr(sum[k]);
                      break;
                      end;
                 end;
        if sumtmp='' then
           tmp:=tmp+_space(l)+' '
        else
           tmp:=tmp+_StrRightSide(sumtmp,l)+' ';

        dash:=dash+_repeatstr('-',l)+' ';
      end;
  writeln(FileHandle,dash);
  writeln(FileHandle,tmp);

  quit:
  TB.GotoBookMark(MyBookMark);
  TB.FreeBookmark(MyBookmark);
  TB.enablecontrols;
  CloseFile(FileHandle);

  _run('notepad.exe '+FileName);

end;
{==============================================================================}
procedure _StrGridDump(SG:TstringGrid ; DisplayWidth:string);
var i,j,l:integer;
    FileHandle : textfile;
    FileName,tmp,dash : string;
    MyBookmark: TBookmark;
    seg:string;
begin
  seg:=' '; {���q�Ÿ�}

  FileName:='C:\Dump.txt';
  if _AllTrim(FileName)='' then exit;

  assignfile(FileHandle,FileName);
  rewrite(FileHandle);


  {�Y}
  if SG.fixedrows<>0 then
     begin
       for j := 1 to SG.fixedrows do
           begin
             tmp:='';
             for i := 0 to (SG.colCount-1) do
                 begin
                   l:=_max(_strtoint( _SegStr(displaywidth,i+1)) , Length(_AllTrim(SG.cells[(i),0])) );
                   if l=0 then l:=10;
                   tmp:=tmp+_restr(SG.cells[(i),j-1],l)+seg;
                 end;
             writeln(FileHandle,tmp);
           end;

       dash:='';
       for i := 0 to (SG.colCount-1) do
           begin
             l:=_max(_strtoint( _SegStr(displaywidth,i+1)) , Length(_AllTrim(SG.cells[(i),0])) );
             if l=0 then l:=10;
             dash:=dash+_repeatstr('=',l)+seg;
           end;
       writeln(FileHandle,dash);
     end;

  {��}
  if (SG.rowCount-SG.fixedrows)<>0 then
     begin
       for j:=SG.fixedrows to (SG.rowCount-1) do
           begin
             tmp:='';
             for i := 0 to (SG.colCount-1) do
                 begin
                   l:=_max(_strtoint( _SegStr(displaywidth,i+1)) , Length(_AllTrim(SG.cells[(i),0])) );
                   if l=0 then l:=10;
                   tmp:=tmp+_restr(SG.cells[(i),j],l)+seg;
                 end;
             writeln(FileHandle,tmp);
           end;
     end;


  CloseFile(FileHandle);

  {showmessage('�ɮ� '+FileName+' ���͡I');}
  _run('notepad.exe '+FileName);

end;
{==============================================================================}
procedure _StructureDump(TB:TTable);
var i : integer;
    FileHandle : textfile;
    FileName,tmp : string;
    fname,ftype,fsize,fbyte,fLength,flabel : string;
    iname,ifields,iexpression : string;
    MyList: TstringList;
begin

  FileName:='C:\Dump.txt';
  if _AllTrim(FileName)='' then exit;

  assignfile(FileHandle,FileName);
  rewrite(FileHandle);

  {
  writeln(FileHandle,'FieldName  '+'DataType      '+'DataBytes '+'DisplayLabel');
  writeln(FileHandle,_repeatstr('-',10)+' '+
                     _repeatstr('-',13)+' '+
                      _repeatstr('-',9)+' '+
                     _repeatstr('-',20));
  for i := 0 to (TB.fieldCount-1) do
    begin
      fname:=_restr(Tb.fields[i].fieldname,10)+' ';
      ftype:=_restr(_SegStr(_FieldType(Tb.fields[i].datatype),1),13)+' ';
      fbyte:=_restr(_SegStr(_FieldType(Tb.fields[i].datatype),2),9)+' ';
      flabel:=_restr(Tb.fields[i].displaylabel,20);
      tmp:=fname+ftype+fbyte+flabel;
      writeln(FileHandle,tmp);
    end; }


  writeln(FileHandle,'< TableName  > : '+TB.tablename);
  writeln(FileHandle,'<DatabaseName> : '+TB.databasename);
  writeln(FileHandle,'DefaultDriver  : '+_AliasDefaultDriver(TB.databasename));
  writeln(FileHandle,'Alias Type     : '+_AliasType(TB.databasename));
    {writeln(FileHandle,'AliasDataBase : '+_AliasDataBase(TB.databasename));}
  writeln(FileHandle,'Alias Path     : '+_AliasDataBasePath(TB.databasename));



  writeln(FileHandle,'');
  writeln(FileHandle,'');

  writeln(FileHandle,'FieldName            '+'DataType      '+'DataSize  ');
  writeln(FileHandle,_repeatstr('-',20)+' '+
                     _repeatstr('-',13)+' '+
                     _repeatstr('-',9));
  TB.FieldDefs.update;
  for i := 0 to (TB.FieldDefs.Count-1) do
    begin
      fname:=_restr(TB.FieldDefs.Items[i].name,20)+' ';   { dbf ��̤j��10}
      ftype:=_restr(_SegStr(_FieldType(TB.FieldDefs.Items[i].datatype),1),13)+' ';
      fsize:=_restr(inttostr(TB.FieldDefs.Items[i].size),9)+' ';
      tmp:=fname+ftype+fsize;
      writeln(FileHandle,tmp);
    end;

  {----------------------}

  writeln(FileHandle,'');
  writeln(FileHandle,'');
  writeln(FileHandle,'IndexName       Fields                         IndexExpression');
  writeln(FileHandle,_repeatstr('-',15)+' '+
                     _repeatstr('-',30)+' '+
                     _repeatstr('-',30));

  TB.IndexDefs.update;
  for i := 0 to (TB.IndexDefs.Count-1) do
    begin
      iname:=_restr(TB.IndexDefs.Items[i].name,15)+' ';
      ifields:=_restr(TB.IndexDefs.Items[i].fields,30)+' ';
      iexpression:=_restr(TB.IndexDefs.Items[I].Expression,30);
      tmp:=iname+ifields+iexpression;
      writeln(FileHandle,tmp);
    end;

  {
  MyList := TstringList.Create;
  Tb.GetIndexNames(MyList);
  for i := 0 to (MyList.Count-1) do
    begin
      writeln(FileHandle,MyList.strings[i]);
    end;
  MyList.Free;}

  CloseFile(FileHandle);

  {showmessage('�ɮ� '+FileName+' ���͡I');}
  _run('notepad.exe '+FileName);

end;
//==============================================================================
// ***** ��ܲ�,������T
//==============================================================================
function _Question(Q:string):Boolean; { ���D��� ��ܲ� }
var I:integer;
begin
I:=application.messagebox(_STRTOPCHAR(Q),'�y�ݤ@�|�� ...',mb_Yesno+mb_defbutton1+
       mb_iconquestion);

if I=6 then Result:=True else Result:=False;

end;
//==============================================================================
function _Input(WhatInput,note,ReadyStr:string):string; { ��J ��ܲ� }
begin
  if inputquery(WhatInput,note,ReadyStr) then
     Result:=ReadyStr
     else
     Result:='';
end;
//==============================================================================
procedure _Show(S:string); { �D�ܹ�ܲ� }
begin
application.messagebox(_STRTOPCHAR(S),'�Ȱ� ...',mb_OK+mb_defbutton1+
       mb_iconstop);
end;
//==============================================================================
procedure _ShowBoolean(B:Boolean);
begin
  if B=True then showmessage('True')
            else showmessage('False');
end;
//==============================================================================
procedure _ShowInt(I:LongInt);
begin
  showmessage(inttostr(I));
end;
//==============================================================================
procedure _ShowFloat(F:Single);
begin
  showmessage(floattostr(F));
end;
{==============================================================================}
procedure _SaveLog(mess:string ; replace:integer);
var LogFile:string;
begin
  LogFile:=_cutright(application.exename,3)+'log';

  if replace=0 then
     _AppendToFile(LogFile,mess) {Append}
  else
     _SaveToFile(LogFile,mess);  {OverWrite}

end;
{==============================================================================}
procedure _ViewLog;
var LogFile:string;
begin
  LogFile:=_cutright(application.exename,3)+'log';
  _notepad(LogFile);
end;
//==============================================================================
// ***** CALL �~���{��
//==============================================================================
procedure _Run(ExeStr:string);
var ExePChar: array[0..255] of Char;
begin
StrPCopy(ExePChar,ExeStr);
WINEXEC(ExePChar,SW_SHOW);
end;
//==============================================================================
procedure _notepad(FileName:string);
begin
  if _file(FileName) then
     _run('notepad.exe '+FileName)
  else
     showmessage(FileName+' not exist !');
end;
//==============================================================================
procedure _emailto(email_address:string);
var email:string;
begin
  email:='mailto:'+email_address;
  shellexecute(0,nil,pchar(email),nil,nil,sw_showdefault);
end;
{==============================================================================}
procedure _wwwto(url:string);
var www:string;
begin
  www:='http://'+url;
  shellexecute(0,nil,pchar(www),nil,nil,sw_showdefault);
end;
//==============================================================================
// ***** �s/�ѽX,�P�_,��T���o
//==============================================================================
function _encrypt(S:string):string; {�r��s/�ѽX}
var i:integer;
    r:string;
begin
   r:='';
   for i := 1 to Length(S) do
      begin
        r:=r+ chr( ord(S[i]) xor (i*7 mod 10) );
      end;

   Result:=r;
end;
{==============================================================================}
function _AppPath:string;
var r:string;
begin
  r:=extractfilepath(application.exename);
  if _right(r,1)<>'\' then r:=r+'\';
  Result:=r;  {�̥k��@�w��\}

  { 1.Getdir(nd,0) -> nd ����
    2.ExtractFileDir(ParamStr(0))  //ParamStr(0) �i�אּ Application.ExeName
    3.ExtractFilepath(ParamStr(0)) //ParamStr(0) �i�אּ Application.ExeName
    //ExtractFileDir �M ExtractFilepath ���P�b��'\' }

end;
{==============================================================================}
function _AppIni:string;   {�t���|}
begin
  Result:=_cutright(application.exename,3)+'ini'; {�Υ� ChangeFileExt(ParamStr(0),'.INI') �]�i}
end;
{==============================================================================}
function _AppName:string;  {�t���|}
begin
  Result:=application.exename;
end;
{==============================================================================}
function _AppExeName:string;  {���t���|,�u���ɦW}
begin
  Result:=extractfilename(application.exename);
end;
{==============================================================================}
function _ChangeExtName(FN,ExtType:string):string; {�N�ɦW�����O�����ɦW ��_ChangeExtName('autoexec.bat','zip') }
begin
  Result:=_CutCharRight(FN,'.',0)+ExtType
end;
{==============================================================================}
function _isEnglishWin:Boolean;
var LocalID:Longint;
begin
  LocalID := GetSystemDefaultLCID;
  if LocalID=$409 then
     Result:=True
  else
     Result:=False;

  // $404 : TAIWAN
  // $804 : CHINA Main_Land
end;
{==============================================================================}
function _isChineseWin:Boolean;
var LocalID:Longint;
begin
  LocalID := GetSystemDefaultLCID;
  if LocalID=$404 then
     Result:=True
  else
     Result:=False;

  // $409 : ENGLISH
  // $804 : CHINA Main_Land
end;
{==============================================================================}
function _isNT : Boolean;
var osv : TOSVERSIONINFO;
begin
  osv.dwosversioninfosize:=sizeof(osv);
  GetVersionEx(osv);
  if osv.dwPlatformId = VER_PLATFORM_WIN32_NT then
     Result := True
  else
     Result := False;
end;
{==============================================================================}
function _myIP:string;
var
  wVersionRequested : WORD;
  wsaData : TWSAData;
  p : PHostEnt;
  s : array[0..128] of char;
  p2 : pchar;
begin
  {Start up WinSock}
  wVersionRequested := MAKEWORD(1, 1);
  WSAStartup(wVersionRequested, wsaData);
  {Get the computer name}
  GetHostName(@s, 128);
  p := GetHostByName(@s);
    //Memo1.Lines.Add(p^.h_Name);
  {Get the IpAddress}
  p2 := iNet_ntoa(PInAddr(p^.h_addr_list^)^);
    //Memo1.Lines.Add(p2);
  {Shut down WinSock}
  WSACleanup;

  Result:=strpas(p2);
end;
{==============================================================================}
function _IsEmail(email:string):Boolean;
var UserName,HostName:string;
begin
  if _SubStrNum(email,'@')<>1 then
     begin
       Result:=False;
       exit;
     end;

  UserName:=_AllTrim(_GetCharSegStr(email,'@',1));
  HostName:=_AllTrim(_GetCharSegStr(email,'@',2));

  if (UserName='') or (HostName='') then
     begin
       Result:=False;
       exit;
     end;

  Result:=True;

end;
{==============================================================================}
function _AutoCaption(English,Chinese:string):string; {�̷Ӥ��^��������� caption}
begin
  if _isChineseWin then
     Result:=chinese
  else
     Result:=english;
end;
{==============================================================================}
function _isLeapYear(yyyy:string):Boolean; {�D�褸�~�r��O�_���|�~}
var y:integer;
    r:Boolean;
begin

  y:=_strtoint(yyyy);

  if y=2000 then
     begin
       Result:=True;
       exit;
     end;

  r:=False;

  if y>2000 then
     begin
       if ((y-2000) mod 4)=0 then r:=True;
     end
  else
     begin
       if ((2000-y) mod 4)=0 then r:=True;
     end;

  Result:=r;

end;
{==============================================================================}
function _isNullTStrings(TS:Tstrings):Boolean; {�P�_�O�_���Ŧr��զ�,�����ťդ]��Ŧr��}
var test,i:integer;
begin
  test:=0;
  for i :=0 to TS.count do
      begin
        if not _isNullStr(TS.strings[i]) then test:=1;
      end;

  if test=0 then Result:=True
  else Result:=False;

end;
{==============================================================================}
function _isNullStr(s:string):Boolean; {�P�_�O�_���Ŧr��,�����ťդ]��Ŧr��}
var str:string;
begin
  str:=s;
  _ReplaceStr(str,' ','');
  _ReplaceStr(str,'�@',''); {�����������ť�}

  if str='' then Result:=True
  else Result:=False;
end;
{==============================================================================}
function _isNumStr(s:string):Boolean;
var i:integer;
begin

  for i:=1 to Length(s) do
      begin
        if not(ord(s[i]) in [48..57]) then
           begin
             Result:=False;
             exit;
           end;
      end;

  Result:=True;
end;
{==============================================================================}
function _isLetter(s:string):Boolean;
var i:integer;
begin

  for i:=1 to Length(s) do
      begin
        if not ((ord(s[i]) in [97..122]) or   {'a'-'z'}
                (ord(s[i]) in [65..90]) or    {'A'-'Z'}
                (ord(s[i]) in [32,46])) then  {' ','.'}
           begin
             Result:=False;
             exit;
           end;
      end;

  Result:=True;
end;
{==============================================================================}
function _isNumKey(key:word):Boolean;
begin

  if key in [48..57] then
     Result:=True
  else
     Result:=False;

end;
{==============================================================================}
function _isEAN13(ean:string):Boolean;
begin
  if (Length(ean)<>13) or (not _isNumStr(ean)) then
     begin
       Result:=False;
       exit;
     end;
  if (copy(ean,1,12)+_EAN13CheakCode(ean))=ean then
     Result:=True
  else
     Result:=False;
end;
{==============================================================================}
function _IdCheck(ID: string): Boolean; { �����Ҹ��X�k���ˬd }
var HEAD:string;
    S:integer;
    WAIT_SEC:SINGLE;
begin
  Result:= False;
  WAIT_SEC:=0.7; {���~�T�����d�ɶ�}

  ID:=_RTRIM(ID);

  if Length(ID)<>10 then
     begin
       {_AutoMessage('�����Ҹ����׿��~!',WAIT_SEC); }
       exit;
     end;

  if not(ID[1] In ['A'..'Z']) then
     begin
       {_AutoMessage('�����Ҹ��Ĥ@�X���~!',WAIT_SEC); }
       exit;
     end;

  if _VALUE(_CUTLeft(ID,1))=0 then
     begin
       {_AutoMessage('�����Ҹ���E�X���~!!',WAIT_SEC);}
       exit;
     end;

   case ID[1] of
        'A':HEAD:='10';
        'B':HEAD:='11';
	'C':HEAD:='12';
        'D':HEAD:='13';
        'E':HEAD:='14';
        'F':HEAD:='15';
        'G':HEAD:='16';
        'H':HEAD:='17';
        'I':HEAD:='34';
        'J':HEAD:='18';
        'K':HEAD:='19';
        'L':HEAD:='20';
        'M':HEAD:='21';
        'N':HEAD:='22';
        'O':HEAD:='35';
        'P':HEAD:='23';
        'Q':HEAD:='24';
        'R':HEAD:='25';
        'S':HEAD:='26';
        'T':HEAD:='27';
        'U':HEAD:='28';
        'V':HEAD:='29';
        'W':HEAD:='32';
        'X':HEAD:='30';
        'Y':HEAD:='31';
        'Z':HEAD:='33';
   end;

	ID:=HEAD+_CUTLeft(ID,1); {�@ 2+9 = 11 ��}

	S:=strtoint(ID[1])+
	   strtoint(ID[2])*9+
	   strtoint(ID[3])*8+
	   strtoint(ID[4])*7+
	   strtoint(ID[5])*6+
	   strtoint(ID[6])*5+
	   strtoint(ID[7])*4+
	   strtoint(ID[8])*3+
	   strtoint(ID[9])*2+
	   strtoint(ID[10]);

	S:=S MOD 10;

	if (S+strtoint(ID[11]))=10 then Result:=True
	   else {_AutoMessage('�����Ҹ��ˬd�X���~!!',WAIT_SEC)} ;

end;
{==============================================================================}
function _GetAge(dDate : TDateTime): integer;	{���o�~��}
var
   nYear1, nYear2, nMonth, nDay: Word;

begin
   if dDate <> 0 then
   begin
      DecodeDate(dDate, nYear1, nMonth, nDay);
      DecodeDate(Date, nYear2, nMonth, nDay);
      Result := nYear2 - nYear1;
   end
   else
       Result := 0;
end;
{==============================================================================}
function _WeekStr(D:TDateTime):string;
begin
  Result:='';
  Case DayofWeek(D) of
       1:Result:='��';
       2:Result:='�@';
       3:Result:='�G';
       4:Result:='�T';
       5:Result:='�|';
       6:Result:='��';
       7:Result:='��';
  end;

end;
{==============================================================================}
function _CountryCode(Tel_Number:string):string;
var telno:string;
    r:string;
    c:integer;
begin
  telno:=Tel_Number;
  telno:=_FilterNumStr(telno);
  if _Left(telno,2)='00' then telno:=_CUTLeft(telno,3);
  if _Left(telno,1)='0' then telno:=_CUTLeft(telno,1);
  if (Length(telno)<=8) and (Length(telno)>=5) then begin Result:='886,Taiwan'; exit; end;

  r:='';
  c:=_strtoint(_Left(telno,4));
  case c of
       1809:  r:='1-809,Grenada'        ;
       1907:  r:='1-907,Alaska (U.S.A.)';
       1268:  r:='1-268,Antigua'        ;
       1242:  r:='1-242,Bahamas'        ;
       1246:  r:='1-246,Barbados'       ;
       1441:  r:='1-441,Bermuda'        ;
       1345:  r:='1-345,Cayman Is'      ;
       1664:  r:='1-664,Montserrat'     ;
       1787:  r:='1-787,Puerto Rico'    ;
       1403,1416,1418,1514,1604,1613,1905:  r:='1,Canada'             ;
       4161,4131,4122:                      r:='41,Switzerland'       ;
  end;

  if r<>'' then
     begin
       Result:=r;
       exit;
     end;

  c:=_strtoint(_Left(telno,3));
  case c of
       212:  r:='212,Morocco'          ;
       213:  r:='213,Algeria'          ;
       216:  r:='216,Tunisia'          ;
       218:  r:='218,Libya'            ;
       220:  r:='220,Gambia'           ;
       221:  r:='221,Senegal'          ;
       222:  r:='222,Mauritania'       ;
       223:  r:='223,Mali Rep'         ;
       224:  r:='224,Guinea'           ;
       225:  r:='225,Ivory Coast'      ;
       226:  r:='226,Burkina Faso'     ;
       227:  r:='227,Niger Rep'        ;
       228:  r:='228,Togo'             ;
       229:  r:='229,Benin'            ;
       230:  r:='230,Mauritius'        ;
       231:  r:='231,Liberia'          ;
       232:  r:='232,Sierra Leone'     ;
       233:  r:='233,Ghana'            ;
       234:  r:='234,Nigeria'          ;
       235:  r:='235,Chad'             ;
       236:  r:='236,Central African Rep'  ;
       237:  r:='237,Cameroon'         ;
       238:  r:='238,Cape Verde Is'    ;
       239:  r:='239,Sao Tome & Principe'  ;
       240:  r:='240,Equatorial Guinea';
       241:  r:='241,Gabon'            ;
       242:  r:='242,Congo'            ;
       243:  r:='243,Zaire'            ;
       244:  r:='244,Angola'           ;
       245:  r:='245,Guinea-Bissau'    ;
       246:  r:='246,Diego Garcia'     ;
       247:  r:='247,Ascension Is'     ;
       248:  r:='248,Seychelles'       ;
       249:  r:='249,Sudan'            ;
       250:  r:='250,Rwanda'           ;
       251:  r:='251,Ethiopia'         ;
       252:  r:='252,Somalia'          ;
       253:  r:='253,Djibouti'         ;
       254:  r:='254,Kenya'            ;
       255:  r:='255,Tanzania'         ;
       256:  r:='256,Uganda'           ;
       257:  r:='257,Burundi'          ;
       258:  r:='258,Mozambique'       ;
       260:  r:='260,Zambia'           ;
       261:  r:='261,Madagascar'       ;
       262:  r:='262,Reunion Is'       ;
       263:  r:='263,Zimbabwe'         ;
       264:  r:='264,Namibia'          ;
       265:  r:='265,Malawi'           ;
       266:  r:='266,Lesotho'          ;
       267:  r:='267,Botswana'         ;
       268:  r:='268,Swaziland'        ;
       269:  r:='269,Mayotte Is'       ;
       290:  r:='290,Saint Helena'     ;
       291:  r:='291,Eritrea'          ;
       297:  r:='297,Aruba'            ;
       298:  r:='298,Faeroe Is'        ;
       299:  r:='299,Greenland'        ;
       350:  r:='350,Gibraltar'        ;
       351:  r:='351,Portugal'         ;
       352:  r:='352,Luxembourg'       ;
       353:  r:='353,Ireland'          ;
       354:  r:='354,Iceland'          ;
       355:  r:='355,Albania'          ;
       356:  r:='356,Malta'            ;
       357:  r:='357,Cyprus'           ;
       358:  r:='358,Finland'          ;
       359:  r:='359,Bulgaria'         ;
       370:  r:='370,Lithuania'        ;
       371:  r:='371,Latvia'           ;
       372:  r:='372,Estonia'          ;
       373:  r:='373,Molodova'         ;

       374:  r:='374,Armenia'          ;
       375:  r:='375,Belarus'          ;
       376:  r:='376,andorra'          ;
       377:  r:='377,Monaco'           ;
       378:  r:='378,San marino'       ;
       380:  r:='380,Ukraine'          ;
       381:  r:='381,Yugoslavia'       ;
       385:  r:='385,Croatia'          ;
       386:  r:='386,Slovenia'         ;
       387:  r:='387,Bosnia-herzegovina'      ;
       389:  r:='389,Macedonia'        ;
       420:  r:='420,Czech Rep'        ;
       421:  r:='421,Slovak'           ;
       500:  r:='500,Falkland Is'      ;
       501:  r:='501,Belize'           ;
       502:  r:='502,Guatemala'        ;
       503:  r:='503,Ei Salvador'      ;
       504:  r:='504,Honduras'         ;
       505:  r:='505,Nicaragua'        ;
       506:  r:='506,Costa Rica'       ;
       507:  r:='507,Panama'           ;
       508:  r:='508,Saint Pierre'     ;
       509:  r:='509,Haiti'            ;
       590:  r:='590,Guadeloupe'       ;
       591:  r:='591,Bolivia'          ;
       592:  r:='592,Guyana'           ;
       593:  r:='593,Ecuador'          ;
       594:  r:='594,French Guiana'    ;
       595:  r:='595,Paraguay'         ;
       596:  r:='596,Martinique Is'    ;
       597:  r:='597,Suriname'         ;
       598:  r:='598,Uruguay'          ;
       599:  r:='599,Netherlands Antilles'       ;
       670:  r:='670,Mariana'          ;
       671:  r:='671,Guam'             ;
       672:  r:='672,Antarctica'       ;
       673:  r:='673,Brunei'           ;
       674:  r:='674,Nauru'            ;
       675:  r:='675,Papua New Guinea' ;
       676:  r:='676,Tonga'            ;
       677:  r:='677,Solomon Is'       ;
       678:  r:='678,Vauatu'           ;
       679:  r:='679,Fiji'             ;
       680:  r:='680,Palau'            ;
       682:  r:='682,Cook Is'          ;

       683:  r:='683,Niue Is'          ;
       684:  r:='684,American Samoa'   ;
       685:  r:='685,Western Samoa'    ;
       686:  r:='686,Kiribati'         ;
       687:  r:='687,New Caledonia'    ;
       688:  r:='688,Tuvalu'           ;
       689:  r:='689,French Polynesia' ;
       691:  r:='691,Micronesia'       ;
       692:  r:='692,Marshall is'      ;
       850:  r:='850,Korea North'      ;
       852:  r:='852,Hong Kong'        ;
       853:  r:='853,Macao'            ;
       855:  r:='855,Cambodia'         ;
       856:  r:='856,Laos'             ;
       880:  r:='880,Bangladesh'       ;
       886:  r:='886,Taiwan'           ;
       960:  r:='960,Maldives'         ;
       961:  r:='961,Lebanon'          ;
       962:  r:='962,Jordan'           ;
       963:  r:='963,Syria'            ;
       964:  r:='964,Iraq'             ;
       965:  r:='965,Kuwait'           ;
       966:  r:='966,Saudi Arabia'     ;
       967:  r:='967,Yemen'            ;
       968:  r:='968,Oman'             ;
       971:  r:='971,U.A.E.'           ;
       972:  r:='972,Israel'           ;
       973:  r:='973,Bahrain'          ;
       974:  r:='974,Qatar'            ;
       975:  r:='975,Bhutan'           ;
       976:  r:='976,Mongolia'         ;
       977:  r:='977,Nepal'            ;
       994:  r:='994,Azerbaijan'       ;
       995:  r:='995,Georgia'          ;
  end;

  if r<>'' then
     begin
       Result:=r;
       exit;
     end;

  c:=_strtoint(_Left(telno,2));
  case c of
       20:   r:='20,Egypt'             ;
       27:   r:='27,South Africa'      ;
       30:   r:='30,Greece'            ;
       31:   r:='31,Netherlands'       ;
       32:   r:='32,Belgium'           ;
       33:   r:='33,France'            ;
       34:   r:='34,Spain'             ;
       36:   r:='36,Hungary'           ;
       39:   r:='39,Italy'             ;
       40:   r:='40,Romania'           ;
       41:   r:='41,Liechtenstein'     ;
       43:   r:='43,Austria'           ;
       44:   r:='44,U.K.'              ;
       45:   r:='45,Denmark'           ;
       46:   r:='46,Sweden'            ;
       47:   r:='47,Norway'            ;
       48:   r:='48,Poland'            ;
       49:   r:='49,Germany'           ;
       51:   r:='51,Peru'              ;
       52:   r:='52,Mexico'            ;
       53:   r:='53,Cuba'              ;
       54:   r:='54,Argentina'         ;
       55:   r:='55,Brazil'            ;
       56:   r:='56,Chile'             ;
       57:   r:='57,Colombia'          ;
       58:   r:='58,Venezuela'         ;
       60:   r:='60,Malaysia'          ;
       61:   r:='61,Australia'         ;
       62:   r:='62,Indonesia'         ;
       63:   r:='63,Phillipines'       ;
       64:   r:='64,New Zealand'       ;
       65:   r:='65,Singapore'         ;
       66:   r:='66,Thailand'          ;
       81:   r:='81,Japan'             ;
       82:   r:='82,Korea South'       ;
       84:   r:='84,Vietnam'           ;
       86:   r:='86,China'             ;
       90:   r:='90,Turkey'            ;
       91:   r:='91,India'             ;
       92:   r:='92,Pakistan'          ;
       93:   r:='93,Afghanistan'       ;
       94:   r:='94,Sri Lanka'         ;
       95:   r:='95,Myanmar'           ;
       98:   r:='98,Iran'              ;
  end;

  if r<>'' then
     begin
       Result:=r;
       exit;
     end;

  c:=_strtoint(_Left(telno,1));
  case c of
       1:    r:='1,U.S.A.'             ;
       7:    r:='7,Russia'             ;
  end;

  if r<>'' then
     begin
       Result:=r;
       exit;
     end;

   Result:='';

end;
//==============================================================================
// ***** API , ����
//==============================================================================
procedure _CloseOtherAP(Windows_Caption:string);
begin
  PostMessage(FindWindow(Nil, pchar(Windows_Caption)), WM_QUIT, 0, 0);
end;
{==============================================================================}
procedure _Delay(ms : longint);
var  TheTime : LongInt;
begin
  TheTime := GetTickCount + ms;
  while GetTickCount < TheTime do
        Application.ProcessMessages;
end;
{==============================================================================}
procedure _Wait(S:Single); { �ɶ�����(59��) �i��J�p�� }
var HOUR,MIN,SEC,MSEC : WORD;
    YEAR,MONTH,DAY : WORD;
    OLD_TIME:TDATETIME;
    SM:SINGLE;
begin
  OLD_TIME:=NOW;
  {DECODEDATE(OLD_TIME,YEAR,MONTH,DAY);}
  DECODETIME((now-OLD_TIME),HOUR,MIN,SEC,MSEC);
  SM:=SEC+(MSEC/1000);

  while SM<S do
	 begin
	   DECODETIME((NOW-OLD_TIME),HOUR,MIN,SEC,MSEC);
	   SM:=SEC+(MSEC/1000);
	 end;
end;
{==============================================================================}
procedure _ForceExec(filename:string);
var execinfo:TSHELLEXECUTEINFO;
begin
   fillchar(execinfo,sizeof(execinfo),0);
   execinfo.cbsize:=sizeof(execinfo);
   execinfo.lpverb:='open';
   execinfo.lpfile:=pchar(filename);
   execinfo.lpparameters:=nil;
   execinfo.fmask:=SEE_MASK_NOCLOSEPROCESS;
   execinfo.nshow:=SW_SHOWDEFAULT;

   shellexecuteex(@execinfo);

   application.minimize;
   waitforsingleobject(execinfo.hprocess,infinite);
   application.restore;
end;
{==============================================================================}
{example : _keyStep(Key,'') ; _keyStep(Key,'no_up') ; _keyStep(Key,'no_next')  }
procedure _keyStep(Key:Word ; option:string);
var o:string;
begin
  o:=lowercase(option);
  if o='no_down' then o:='no_next';
  if (key=VK_UP) and (o<>'no_up') then _up;
  if ((key=VK_down) or (key=VK_RETURN)) and (o<>'no_next') then _next;
end;
{==============================================================================}
procedure _Next; {�۰ʱN�ʧ@���� forM ���J�I���ܨ䤺�U�@����}
begin
  POSTMESSAGE(SCREEN.ACTIVEFORM.HandLE,WM_NEXTDLGCTL,0,0);

end;
{==============================================================================}
procedure _UP; {�۰ʱN�ʧ@���� forM ���J�I���ܨ䤺�W�@����}
begin
   POSTMESSAGE(SCREEN.ACTIVEFORM.HandLE,WM_NEXTDLGCTL,1,0);
end;
{==============================================================================}
procedure _ReSetFocus;
begin
  _next;
  _up;
end;
{==============================================================================}
function _Resource(GU:string):string; {���o�Ѿl�귽��,�ǤJ�ѼƦ�'GDI'��'USER'}
var I:integer;
begin
  {$IFDEF WIN32}
    showmessage('_RESOURCE ��Ƥ��䴩 32 �줸�t��');
    Result:='';
  {$else}
  GU:=LOWERCASE(GU);

  if GU='system' then
     I:=GETFREESYSTEMRESOURCES(GFSR_SYSTEMRESOURCES);

  if GU='gdi' then
     I:=GETFREESYSTEMRESOURCES(GFSR_GDIRESOURCES);

  if GU='user' then
     I:=GETFREESYSTEMRESOURCES(GFSR_USERRESOURCES);

  if (GU<>'system') and (GU<>'gdi') and (GU<>'user') then
     begin
       showmessage('_RESOURCE() �ǤJ�Ѽƿ��~');
       Result:='';
       exit;
     end;

  Result:=(inttostr(I)+' %');
  {$endIF}

end;
{==============================================================================}
procedure _CapLock(bLockIt: Boolean); { ��L�j�p�g��w }
Var
  Level : integer;		  { _CapLock(True) ; �j�g }
  KeyState : TKeyBoardState;	  { _CapLock(False) ; �p�g }
begin
  Level := GetKeyState(VK_CAPITAL);
  GetKeyboardState(KeyState);
  if bLockIt then
    KeyState[VK_CAPITAL] := 1
  else
    KeyState[VK_CAPITAL] := 0;
  setKeyboardState(KeyState);
end;
//==============================================================================
// ***** DateTime ����ɶ��B�z
//==============================================================================
function _WhatTime(What:string):string; { ����{�b�ɶ�����,��,�� }
var Hour,Min,Sec,MSec : Word;		{ �ǤJ�ѼƦ� 'hour','min','sec','time' }
    H,M,S,MS : string;			{ �� _WHATTIME('hour') �Ǧ^�T�w��줧�p�ɦr��}
begin					{    _WHATTIME('min') �Ǧ^�T�w��줧�����r��}
                                        {    _WHATTIME('sec') �Ǧ^�T�w��줧��r��}
  DeCodeTime(Time,Hour,Min,Sec,MSec);   {    _WHATTIME('time') �Ǧ^�T�w�K�줧'HH:MM:SS'�ɶ��r��}
  H:=inttostr(Hour);                    {    _WHATTIME('hourmin') �Ǧ^�T�w���줧'HH:MM'�ɶ��r��}
     if Length(H)=1 then H:='0'+H;      {    _WHATTIME('msec') �Ǧ^�T�w�T�줧�@��r��}
  M:=inttostr(Min);
     if Length(M)=1 then M:='0'+M;
  S:=inttostr(Sec);
     if Length(S)=1 then S:='0'+S;

  What:=LowerCase(What);
  Result:='';
  if What='hour' then Result:=H;
  if What='min' then Result:=M;
  if What='sec' then Result:=S;
  if What='time' then Result:=H+':'+M+':'+S;
  if What='hourmin' then Result:=H+':'+M;
  if WHAT='msec' then Result:=_STRZERO(msec,3);  {�r���000-999}

end;
{==============================================================================}
function _WhatDate(What:string):string; { ����{�b������~,��,�� }
var YEAR,MONTH,DAY : WORD;		{ �ǤJ�ѼƦ� 'year','rocyear,'month','day','date','rocdate' }
    Y,CY,M,D,LONGY : string;		{ �� _WHATDATE('year') �Ǧ^�T�w��줧�褸�~�r��}
begin					{    _WHATDATE('rocyear') �Ǧ^�T�w��줧����~�r��}
DECODEDATE(DATE,YEAR,MONTH,DAY);	{    _WHATDATE('month') �Ǧ^�T�w��줧��r��}
Y:=inttostr(YEAR-1900); 		{    _WHATDATE('day') �Ǧ^�T�w��줧��r��}
   if Length(Y)=1 then Y:='0'+Y;        {    _WHATDATE('date') �Ǧ^�T�w�K�줧'MM.DD.YY'�褸����r��(�~�b�̫�)}
CY:=inttostr(YEAR-1911);		{    _WHATDATE('rocdate') �Ǧ^�T�w�K�줧'yy.MM.DD'�������r��}
   if Length(CY)=1 then CY:='0'+CY;     {    _WHATDATE('lyear') �Ǧ^�T�w�|�줧�褸�~�r��}
M:=inttostr(MONTH);
   if Length(M)=1 then M:='0'+M;
D:=inttostr(DAY);
   if Length(D)=1 then D:='0'+D;
LONGY:=inttostr(YEAR);


WHAT:=LOWERCASE(WHAT);
Result:='';
if WHAT='year' then Result:=Y;
if WHAT='lyear' then Result:=LONGY;
if WHAT='rocyear' then Result:=CY;
if WHAT='month' then Result:=M;
if WHAT='day' then Result:=D;
if WHAT='date' then Result:=M+'.'+D+'.'+Y;      {�~�b�᭱}
if WHAT='rocdate' then Result:=CY+'.'+M+'.'+D;  {�~�b�e��}

end;
{==============================================================================}
{�D�ǤJ�~�뤧�U�@�~��}
function _AddYearMonth(yymm:string ; inc:integer):string;
var y,m,l,i:integer;
begin

if (Length(_AllTrim(yymm))<> 4) and (Length(_AllTrim(yymm))<> 6) then
   begin
     showmessage('�~��榡���~�I( '+yymm+' )');
     Result:=yymm;
     exit;
   end;

y:=strtoint(_cutright(yymm,2));
m:=strtoint(_right(yymm,2));
l:=Length(inttostr(y));

//if (y=99) and (m=12) then exit;

if inc<=0 then
   begin
     Result:=yymm;
     exit;
   end;

for i:=1 to inc do
  begin
     m:=m+1;
     if m=13 then
        begin
          m:=1;
          y:=y+1;
        end;
  end;

Result:=_strzero(y,l)+_strzero(m,2);

end;
{==============================================================================}
{�D�ǤJ�~�뤧�e�@�~��}
function _DecYearMonth(yymm:string ; dec:integer):string;
var y,m,l,i:integer;
begin

if (Length(_AllTrim(yymm))<> 4) and (Length(_AllTrim(yymm))<> 6) then
   begin
     showmessage('�~��榡���~�I( '+yymm+' )');
     Result:=yymm;
     exit;
   end;

y:=strtoint(_cutright(yymm,2));
m:=strtoint(_right(yymm,2));
l:=Length(inttostr(y));

if dec<=0 then
   begin
     Result:=yymm;
     exit;
   end;

for i:=1 to dec do
  begin
    m:=m-1;
    if m=0 then
       begin
         m:=12;
         y:=y-1;
       end;
  end;

Result:=_strzero(y,l)+_strzero(m,2);

end;
{==============================================================================}
{�D�{�b�~��}
function _ThisYearMonth(f:integer):string;
begin

if f=0 then
   Result:=_whatdate('lyear')+_whatdate('month')   {�覡�~��}
else
   Result:=_whatdate('rocyear')+_whatdate('month'); {�����~��}

end;
{==============================================================================}
{�ɶ��r��ۥ[ , ��q�T�ɶ��O�ήɥ�}
function _AddTimeStr(TimeStr1,TimeStr2:string):string;
var h,m,s:integer;
    c:integer;
begin
  h:=0;m:=0;s:=0;
  c:=0;

  s:=_strtoint(_GetCharSegStr(TimeStr1,':',3))+_strtoint(_GetCharSegStr(TimeStr2,':',3));
  c:=_int(s/60);
  s:=(s mod 60);

  m:=_strtoint(_GetCharSegStr(TimeStr1,':',2))+_strtoint(_GetCharSegStr(TimeStr2,':',2))+c;
  c:=_int(m/60);
  m:=(m mod 60);

  h:=_strtoint(_GetCharSegStr(TimeStr1,':',1))+_strtoint(_GetCharSegStr(TimeStr2,':',1))+c;

  if h>99 then
     Result:=inttostr(h)+':'+_strzero(m,2)+':'+_strzero(s,2)
  else
     Result:=_strzero(h,2)+':'+_strzero(m,2)+':'+_strzero(s,2);

end;
{==============================================================================}
function _TimeToMinute(TimeStr:string):string; {�ɶ��ন����}
var h,m,s:integer;
    c:integer;
begin
  h:=0;m:=0;s:=0;
  c:=0;

  s:=_strtoint(_GetCharSegStr(TimeStr,':',3));
  m:=_strtoint(_GetCharSegStr(TimeStr,':',2));
  h:=_strtoint(_GetCharSegStr(TimeStr,':',1));

  m:=(m+(h*60));

  Result:=inttostr(m)+':'+_strzero(s,2)

end;
{==============================================================================}
function _TimeToSecond(TimeStr:string):string; {�ɶ��ন����}
var h,m,s:integer;
    c:integer;
begin
  h:=0;m:=0;s:=0;
  c:=0;

  s:=_strtoint(_GetCharSegStr(TimeStr,':',3));
  m:=_strtoint(_GetCharSegStr(TimeStr,':',2));
  h:=_strtoint(_GetCharSegStr(TimeStr,':',1));

  s:=s+(m*60)+(h*3600);

  Result:=inttostr(s);

end;
{==============================================================================}
function _nowstr:string; {�s�ɮװߤ@�W�r�ɨ��W�� 19990703141000}
begin
  Result:=_whatdate('lyear')+_whatdate('month')+_whatdate('day')+
          _whattime('hour')+_whattime('min')+_whattime('sec');
end;
{==============================================================================}
function _now:string; {�D�{�b����ɶ�(�褸�~) 1999/07/03 14:10:00}
begin
  Result:=_whatdate('lyear')+'/'+_whatdate('month')+'/'+_whatdate('day')+' '+
          _whattime('hour')+':'+_whattime('min')+':'+_whattime('sec');
end;
{==============================================================================}
function _now2:string; {�D�{�b����ɶ�(�褸�~) 1999.07.03 14:10:00}
begin
  Result:=_whatdate('lyear')+'.'+_whatdate('month')+'.'+_whatdate('day')+' '+
          _whattime('hour')+':'+_whattime('min')+':'+_whattime('sec');
end;
{==============================================================================}
{�D�褸�~�뤧�Ӥ�̫�@�Ѥ���r�� ; onlyday=1 ��u�Ǧ^�Ӥ�ѼƦr��}
function _monthLastDay(yyyymm:string ; onlyday:integer):string;
var y:string;
    m:integer;
    d:string;
begin

  if Length(yyyymm)<>6 then
     begin
       Result:='';
       exit;
     end;

  y:=_Left(yyyymm,4);
  m:=_strtoint(_right(yyyymm,2));

  d:='';
  case m of
     1,3,5,7,8,10,12 : d:='31';
     4,6,9,11	     : d:='30';
     2		     : if _isLeapYear(y) then d:='29' else d:='28';
  end;

  if d='' then
     begin
       Result:='';
       exit;
     end;

 if onlyday=0 then
    Result:=yyyymm+d
 else
    Result:=d;


end;
{==============================================================================}
{i:=0 is today ; i:=-1 is yesterday ; i:=1 is tomorrow ; example _date(0,'/') }
function _date(i:integer ; c:string):string; {1999/07/03 or 1999.07.03 or 1999-07-03}
var Present: TDateTime;
    Year, Month, Day : Word;
begin
  Present:=date+i;
  DecodeDate(Present, Year, Month, Day);
  Result:=_strzero(Year,4)+c+_strzero(Month,2)+c+_strzero(Day,2);
end;
{==============================================================================}
function _dateToStr(D:TDatetime ; c:string):string; {1999/07/03 or 1999.07.03 or 1999-07-03}
var Year, Month, Day : Word;
begin
  DecodeDate(D, Year, Month, Day);
  Result:=_strzero(Year,4)+c+_strzero(Month,2)+c+_strzero(Day,2);
end;
{==============================================================================}
{i:=0 is today ; i:=-1 is yesterday ; i:=1 is tomorrow }
function _dateToYYYYMMDD(i:integer):string; {19990703}
var Present: TDateTime;
    Year, Month, Day : Word;
begin
  Present:=date+i;
  DecodeDate(Present, Year, Month, Day);
  Result:=_strzero(Year,4)+_strzero(Month,2)+_strzero(Day,2);
end;
{==============================================================================}
function _dateToMMDD(i:integer):string; {0703}
var Present: TDateTime;
    Year, Month, Day : Word;
begin
  Present:=date+i;
  DecodeDate(Present, Year, Month, Day);
  Result:=_strzero(Month,2)+_strzero(Day,2);
end;
//==============================================================================
// ***** ROC DateTime �������ɶ��B�z
//==============================================================================
function _ROCDATE(DD:TDATETIME;P:integer):string; {�ഫ�Y���������0YYMMDD �����r�� }
var YEAR,MONTH,DAY : WORD;			  {P=0 ���['.'}
    Y,CY,M,D,LONGY : string;			  {P=1 �['.' }
    YY:integer;
begin
DECODEDATE(DD,YEAR,MONTH,DAY);

if (year=0) and (month=0) and (day=0) then
   begin
     Result:='';
     exit;
   end;

YY:=YEAR-1911;
if YY>0 then
   begin
     CY:=inttostr(YY);
      if Length(CY)=1 then CY:='00'+CY;
      if Length(CY)=2 then CY:='0'+CY;
   end
else
   begin
   YY:=YEAR-1912;
   CY:=inttostr(YY);
      if Length(CY)=2 then CY:='-0'+_RIGHT(CY,1);
   end;

if strtoint(CY)>999 then
   CY:='XXX';

if (CY<>'XXX') and (strtoint(CY)<-99) then
   CY:='-XX';

M:=inttostr(MONTH);
   if Length(M)=1 then M:='0'+M;
D:=inttostr(DAY);
   if Length(D)=1 then D:='0'+D;

if P=0 then
 Result:=CY+M+D
else
 Result:=CY+'.'+M+'.'+D;

end;
{==============================================================================}
function _ROCCHECK(DD:string):Boolean;
var ROCYEAR,M,D : string;
    F:Boolean;
    L:integer;
    MM:integer;
begin
F:=False;
DD:=_RTRIM(DD);
DD:=_LTRIM(DD);

if POS('.',DD)>0 then
   begin
     if (Length(DD)<>8) and (Length(DD)<>9) then
	begin
	Result:=F;
	{_SHOW('�������r����צ��~');}
	exit;
	end;
   end
else
   begin
     if (Length(DD)<>6) and (Length(DD)<>7) then
	begin
	Result:=F;
	{_SHOW('�������r����צ��~');}
	exit;
	end;
   end;

if (Length(DD)=6) or (Length(DD)=7) then
   begin
   M:=_RCOPY(DD,4,2);
      if (strtoint(M)>12) or (strtoint(M)<1) then
	  begin
	  {_SHOW('�������r�ꤧ������~');}
	  Result:=F;
	  exit;
	  end;
   ROCYEAR:= _CUTRIGHT(DD,4);
      L:=((strtoint(ROCYEAR)+1911) MOD 4);
   D:=_RIGHT(DD,2);
      MM:=strtoint(M);
      CASE MM OF
	   1,3,5,7,8,10,12:if (strtoint(D)>31) or (strtoint(D)<1) then
				begin
				{_SHOW('�������r�ꤧ������~');}
				Result:=F;
				exit;
				end;
	   2:		   begin
			   if (L=0) and ((strtoint(D)>29) or (strtoint(D)<1)) then
				begin
				{_SHOW('�������r�ꤧ������~');}
				Result:=F;
				exit;
				end;
			   if (L<>0) and ((strtoint(D)>28) or (strtoint(D)<1)) then
				begin
				{_SHOW('�������r�ꤧ������~');}
				Result:=F;
				exit;
				end;
			   end;
	   4,6,9,11:	   if (strtoint(D)>30) or (strtoint(D)<1) then
				begin
				{_SHOW('�������r�ꤧ������~');}
				Result:=F;
				exit;
				end;
      end;
   end;

if (Length(DD)=8) or (Length(DD)=9) then
   begin

   if _SubStrNum(DD,'.')<>2 then
      begin
	{_SHOW('�������r�ꤧ�I�Ʀ��~');}
	Result:=F;
	exit;
      end;
   if (_RCOPY(DD,3,1)<>'.') or (_RCOPY(DD,6,1)<>'.') then
      begin
	{_SHOW('�������r�ꤧ�I��m���~'); }
	Result:=F;
	exit;
      end;

   M:=_RCOPY(DD,5,2);
      if (strtoint(M)>12) or (strtoint(M)<1) then
	  begin
	  {_SHOW('�������r�ꤧ������~');}
	  Result:=F;
	  exit;
	  end;
   ROCYEAR:= _CUTRIGHT(DD,6);
      L:=((strtoint(ROCYEAR)+1911) MOD 4);
   D:=_RIGHT(DD,2);
      MM:=strtoint(M);
      CASE MM OF
	   1,3,5,7,8,10,12:if (strtoint(D)>31) or (strtoint(D)<1) then
				begin
				{_SHOW('�������r�ꤧ������~');}
				Result:=F;
				exit;
				end;
	   2:		   begin
			   if (L=0) and ((strtoint(D)>29) or (strtoint(D)<1)) then
				begin
				{_SHOW('�������r�ꤧ������~');}
				Result:=F;
				exit;
				end;
			   if (L<>0) and ((strtoint(D)>28) or (strtoint(D)<1)) then
				begin
				{_SHOW('�������r�ꤧ������~');}
				Result:=F;
				exit;
				end;
			   end;
	   4,6,9,11:	   if (strtoint(D)>30) or (strtoint(D)<1) then
				begin
				{_SHOW('�������r�ꤧ������~');}
				Result:=F;
				exit;
				end;
      end;
   end;

F:=True;
Result:=F;
end;
{==============================================================================}
function _ROCADD(DD:string;N,P:integer):string;
var Y,M,D:string;
begin
  if _ROCCHECK(DD)=False then
     begin
     showmessage('_ROCADD<> ��ƿ�J�Ѽƿ��~�I');
     Result:='';
     exit;
     end;


  D:=_RIGHT(DD,2);
  DD:=_CUTRIGHT(DD,2);
  if _RIGHT(DD,1)='.' then DD:=_CUTRIGHT(DD,1);
  M:=_RIGHT(DD,2);
  DD:=_CUTRIGHT(DD,2);
  if _RIGHT(DD,1)='.' then DD:=_CUTRIGHT(DD,1);
  Y:=inttostr(strtoint(DD)+1911);

  {$IFDEF WIN32}
  Result:=_ROCDATE((STRTODATE(Y+'/'+M+'/'+D)+N),P);    {DELPHI 2.01}
  {$else}
  Result:=_ROCDATE((STRTODATE(M+'/'+D+'/'+Y)+N),P);    {DELPHI 1.0}
  {$endIF}

end;
{==============================================================================}
function _ROCDEC(DD1,DD2:string):LONGINT;   {�DDD1-DD2 ����������t}
var Y,M,D:string;
    Y2,M2,D2:string;
    RR:string;
begin
  if (_ROCCHECK(DD2)=False) or (_ROCCHECK(DD1)=False) then
     begin
     showmessage('_ROCDEC<> ��ƿ�J�Ѽƿ��~�I');
     Result:=0;
     exit;
     end;

  D:=_RIGHT(DD2,2);
  DD2:=_CUTRIGHT(DD2,2);
  if _RIGHT(DD2,1)='.' then DD2:=_CUTRIGHT(DD2,1);
  M:=_RIGHT(DD2,2);
  DD2:=_CUTRIGHT(DD2,2);
  if _RIGHT(DD2,1)='.' then DD2:=_CUTRIGHT(DD2,1);
  Y:=inttostr(strtoint(DD2)+1911);

  D2:=_RIGHT(DD1,2);
  DD1:=_CUTRIGHT(DD1,2);
  if _RIGHT(DD1,1)='.' then DD1:=_CUTRIGHT(DD1,1);
  M2:=_RIGHT(DD1,2);
  DD1:=_CUTRIGHT(DD1,2);
  if _RIGHT(DD1,1)='.' then DD1:=_CUTRIGHT(DD1,1);
  Y2:=inttostr(strtoint(DD1)+1911);

  {$IFDEF WIN32}
  RR:=FLOATTOSTR((STRTODATE(Y2+'/'+M2+'/'+D2))-(STRTODATE(Y+'/'+M+'/'+D))); {DELPHI 2.01}
  {$else}
  RR:=FLOATTOSTR((STRTODATE(M2+'/'+D2+'/'+Y2))-(STRTODATE(M+'/'+D+'/'+Y))); {DELPHI 1.0}
  {$endIF}

  Result:=strtoint(RR);

end;
{==============================================================================}
function _RocNow(T:integer):string;
begin
CASE T OF
  0: Result:='�{�b����G'+ _whatdate('rocdate')+'�@�{�b�ɶ��G'+_whattime('time');
  1: Result:='�{�b����G'+ _whatdate('rocdate')+'�@�{�b�ɶ��G'+_whattime('hourmin');
  2: Result:='����G'+ _whatdate('rocdate')+'�@�ɶ��G'+_whattime('time');
  3: Result:='����G'+ _whatdate('rocdate')+'�@�ɶ��G'+_whattime('hourmin');
end;

end;
{==============================================================================}
function _RocReStr(DD:string;T:integer):string; {�������r��(����������T���ˬd)}
var Y,M,D:string;
begin
  if (_RTRIM(DD)='') or (T<0) then
     begin
       Result:=DD;
       exit;
     end;

  if (_SubStrNum(DD,'.')<>0) and (_SubStrNum(DD,'.')<>2) then
     begin
       {_SHOW('����r��L�k����έ���I(�`�I�ƿ��~)');}
       Result:=DD;
       exit;
     end;

  DD:=_CutAllChar(DD,' ');

  Y:=_ROCGetSeg(DD,1);
  M:=_ROCGetSeg(DD,2);
  D:=_ROCGetSeg(DD,3);

  Y:=_FilterNumStr(Y);
  M:=_FilterNumStr(M);
  D:=_FilterNumStr(D);

  M:=_CutAllChar(M,'-');  {��Τ�L�t��}
  D:=_CutAllChar(D,'-');

  if (Length(Y)>3) or (Length(M)>2) or (Length(D)>2) then
     begin
       {_SHOW('����r��L�k����έ���I');}
       Result:=DD;
       exit;
     end;

  if Length(D)=0 then D:='  ';
  if Length(D)=1 then D:='0'+D;

  if Length(M)=0 then M:='  ';
  if Length(M)=1 then M:='0'+M;

  if Length(Y)=0 then Y:='   ';
  if Length(Y)=1 then Y:='00'+Y;
  if Length(Y)=2 then Y:='0'+Y;
  Y:=_FilterNumStr(Y);	{�ɧU��N�t�����e���\��}

  CASE T OF
    0: Result:=Y+M+D;
    1: Result:=Y+'.'+M+'.'+D;
  end;

end;
{==============================================================================}
function _ROCCheckF(ROC:string;F:integer):Boolean;
var Y,M,D : string;
    YY,MM,DD : integer;
    PN:integer;
begin
  PN:=_SubStrNum(ROC,'.');

  if (PN<>0) and (PN<>2) then
     begin
       Result:=False;
       exit;
     end;

  if (F=1) and (
     (Length(ROC)<6) or (Length(ROC)>9)  ) then     {���׮榡�ˬd}
     begin
       Result:=False;
       exit;
     end;

  if PN=0 then
     begin
       D:=_RIGHT(ROC,2);
       M:=_RCOPY(ROC,4,2);
       Y:=_CUTRIGHT(ROC,4);
     end;

  if PN=2 then
     begin
       Y:=_GetCharSegStr(ROC,'.',1);
       M:=_GetCharSegStr(ROC,'.',2);
       D:=_GetCharSegStr(ROC,'.',3);
     end;

  if (_strtoint(Y)=0) OR
     (_strtoint(M)=0) OR
     (_strtoint(D)=0) then
     begin
       Result:=False;
       exit;
     end;

  if (F=1) and (
     ((Length(Y)<>2) and (Length(Y)<>3)) OR	 {���׮榡�ˬd}
     (Length(M)<>2) OR
     (Length(D)<>2) ) then
     begin
       Result:=False;
       exit;
     end;

  YY:=_strtoint(Y);
  MM:=_strtoint(M);
  DD:=_strtoint(D);

  if ((YY<-99) or (YY>999)) OR
     ((MM<1) or (MM>12)) OR
     ((DD<1) or (DD>31)) then
     begin
       Result:=False;
       exit;
     end;

  if DD<=_ROCMonthDay(Y,M) then Result:=True
			     else Result:=False;

end;
{==============================================================================}
function _ROCMonthDay(ROCY,MM:string):integer;
var M : integer;
begin
M:=_strtoint(MM);

if _strtoint(ROCY)=0 then
   begin
     Result:=0;
     exit;
   end;

CASE M OF
     1,3,5,7,8,10,12 : Result:=31;
     4,6,9,11	     : Result:=30;
     2		     : if _ROCLYearCheck(ROCY) then Result:=29
					       else Result:=28;
else
     Result:=0;
end;

end;
{==============================================================================}
function _ROCLYearCheck(ROCY:string):Boolean;
var RY,Y,L:integer;
begin
  RY:=_strtoint(ROCY);

  if RY=0 then
     begin
     Result:=False;
     exit;
     end;

  if RY>0 then Y:=RY+1911;
  if RY<0 then Y:=RY+1912;

  L:=(Y MOD 4);

  if L=0 then Result:=True
     else Result:=False;

end;
{==============================================================================}
function _ROCGetSeg(ROC:string;Seg:integer):string; {��������r�ꤤ���~���(�������T���ˬd)}
var PN:integer;
    Y,M,D : string;
begin
  PN:=_SubStrNum(ROC,'.');

  if (PN<>0) and (PN<>2) then
     begin
       Result:='';
       exit;
     end;

  if PN=0 then
     begin
       D:=_RIGHT(ROC,2);
       M:=_RCOPY(ROC,4,2);
       Y:=_CUTRIGHT(ROC,4);
     end;

  if PN=2 then
     begin
       Y:=_GetCharSegStr(ROC,'.',1);
       M:=_GetCharSegStr(ROC,'.',2);
       D:=_GetCharSegStr(ROC,'.',3);
     end;

  CASE Seg OF
       1 : begin
	     {if ((Length(Y)<>2) and (Length(Y)<>3)) OR
		(_strtoint(Y)=0) then
		Result:='' else }
		Result:=Y;
	   end;
       2 : begin
	     {if (Length(M)<>2) OR
		((_strtoint(M)<1) or (_strtoint(M)>12)) then
		Result:='' else  }
		Result:=M;
	   end;
       3 : begin
	     {if (Length(D)<>2) OR
		((_strtoint(D)<1) or (_strtoint(D)>31)) then
		Result:='' else  }
		Result:=D;
	   end;
  else
    Result:='';
  end;
end;
{==============================================================================}
function _RocToDate(Roc:string):TDateTime;
var Y,M,D : string;
begin
  Result:=StrToDate('01/01/01');

  if _RocCheckF(ROC,1) then
     begin
       Y:=inttostr(_strtoint(_ROCGETSEG(ROC,1))+1911);
       M:=_ROCGETSEG(ROC,2);
       D:=_ROCGETSEG(ROC,3);
       {$IFDEF WIN32}
       Result:=STRTODATE(Y+'/'+M+'/'+D);    {DELPHI 2.01}
       {$else}
       Result:=STRTODATE(M+'/'+D+'/'+Y);    {DELPHI 1.0}
       {$endIF}
     end;

end;
{==============================================================================}
function _ROCMonthFirstDay(ROCY,M:string;P:integer):string;
begin
  Result:='';
  if (_strtoint(ROCY)=0) or (_strtoint(M)=0) then exit;
  if not(_strtoint(M) in [1..12]) then exit;

  ROCY:=_strzero(_strtoint(ROCY),3);
  M:=_strzero(_strtoint(M),2);

  if P=0 then Result:=(ROCY+M+'01') else
              Result:=(ROCY+'.'+M+'.01');

end;
{==============================================================================}
function _ROCMonthLastDay(ROCY,M:string;P:integer):string;
var LastDay:string;
begin
  Result:='';
  if (_strtoint(ROCY)=0) or (_strtoint(M)=0) then exit;
  if not(_strtoint(M) in [1..12]) then exit;


  ROCY:=_strzero(_strtoint(ROCY),3);
  M:=_strzero(_strtoint(M),2);
  LastDay:=_strzero(_ROCMONTHDAY(ROCY,M),2);

  if P=0 then Result:=(ROCY+M+LastDay) else
              Result:=(ROCY+'.'+M+'.'+LastDay);

end;
//==============================================================================
// ***** �q�T�O�v , �r��B�I�B�z
//==============================================================================
{�D�����}
function _BaseNumber(Son,Mon:LongInt):LongInt;
var r:LongInt;
begin
r:=0;
if (Son<>0) and (Mon<>0) then
   begin
     r:=(_int((Son-1)/Mon)+1);
   end;

Result:=r;

end;
{==============================================================================}
{�����T�w�p�Ʀ줧�ƭȦr�� --  ���->�r��  }
function _str(f:single ; digits:integer):string;
begin
Result:=formatfloat(('0.'+_RepeatStr('0',digits)),f);
end;
{==============================================================================}
{�����T�w�p�Ʀ줧�ƭȦr�� --  �r��->�r��  }
function _string(f:string ; digits:integer):string;
begin
  if _AllTrim(f)='' then f:='0';
  Result:=formatfloat(('0.'+_RepeatStr('0',digits)),strtofloat(f));
end;
{==============================================================================}
{�D�q�H�O�v;��ƶǤJ,��ƶǥX}
function _cost(TimeBase:integer ; BaseRate,ExtenRate:single):single;
var r:single;
begin
  r:=0;
  if TimeBase<>0 then
     begin
       r:=BaseRate;
       r:=(r+(TimeBase-1)*ExtenRate);
     end;
  Result:=r;
end;
{==============================================================================}
{�D�q�H�O�v;�r��ǤJ,�r��ǥX(�@��p�Ʈ榡)--> ������� }
function _cost2(TimeBase:integer ; BaseRate,ExtenRate:string):string;
var r:single;
begin
  r:=0;
  if TimeBase<>0 then
     begin
       r:=strtofloat(BaseRate);
       r:=(r+(TimeBase-1)*strtofloat(ExtenRate));
     end;
  Result:=_str(r,1);
end;
{==============================================================================}
{�D�q�H�O�v}
function _charge(TimeBase , BaseTime , ExtenTime , BaseRate , ExtenRate:single):single;
var r:single;
    TB,BT,ET:integer;
begin
  if (TimeBase=0) or (BaseTime=0) or (ExtenTime=0) or (BaseRate=0) or (ExtenRate=0) then
     begin
       Result:=0;
       exit;
     end;

  r:=BaseRate;

  if TimeBase>BaseTime then
     begin
       //��1000�i��T��p��2��,��TimeBase�ӷ��i�ܤp��2��)
       TB:=_int(TimeBase*1000);  {�ܾ�Ƥ~���|���B�I�~�t}
       BT:=_int(BaseTime*1000);
       ET:=_int(ExtenTime*1000);
       r:=r+(_Mininteger((TB-BT)/ET)*ExtenRate);
     end;

  Result:=r;

end;
//==============================================================================
// ***** EAN_13
//==============================================================================
function _EAN13CheakCode(ean:string):string;
var sum:integer;
begin

  if (Length(ean)<12) or (not _isNumStr(ean)) then
     begin
       Result:='';
       exit;
     end;

  sum:=strtoint(copy(ean,1,1))+strtoint(copy(ean,3,1))+strtoint(copy(ean,5,1))+
       strtoint(copy(ean,7,1))+strtoint(copy(ean,9,1))+strtoint(copy(ean,11,1))+
       strtoint(copy(ean,2,1))*3+strtoint(copy(ean,4,1))*3+strtoint(copy(ean,6,1))*3+
       strtoint(copy(ean,8,1))*3+strtoint(copy(ean,10,1))*3+strtoint(copy(ean,12,1))*3;
  //showmessage(inttostr(sum));
  Result:=inttostr((10-(sum mod 10)) mod 10); {if 10 then 0}

end;
{==============================================================================}
function _toBeEAN13(ean:string):string;
var sum:integer;
begin

  if (Length(ean)<12) or (not _isNumStr(ean)) then
     begin
       Result:='';
       exit;
     end;

  Result:=copy(ean,1,12)+_EAN13CheakCode(ean);

end;
{==============================================================================}
function _serialNumAdd(serial:string):string;
var i:integer;
    sl:longint;
    r,s:string;
begin

  s:='';
  for i:= 1 to Length(serial) do
  begin
    r:=_right(serial,i);
    if (not _isNumStr(r)) or (i>9) then break;
    s:=r;
  end;

  if s<>'' then
     begin
       sl:=strtoint(s)+1;
       s:=_strzero(sl,Length(s));
       Result:=_Left(serial,Length(serial)-Length(s))+s;
     end
  else Result:=serial;

end;
{==============================================================================}
function _serialNumAddEAN13(ean13:string):string;
var serial:string;
begin

  if (Length(ean13)<12) or (not _isNumStr(ean13)) then
     begin
       showmessage('not an EAN_13 code !');
       Result:='';
       exit;
     end;

  serial:=_serialNumAdd(_Left(ean13,12));
  Result:=_toBeEAN13(serial);

end;
//==============================================================================
// ***** DATABASE ��Ʈw
//==============================================================================
{ IndexName='' �ϥέ���� }
function _TableFind(TB:TTable ; IndexName:string ; parameter:string):Boolean;
begin
  if IndexName<>'' then TB.indexname:=IndexName;
  TB.refresh;
  Result:=TB.findkey([parameter]);
end;
{==============================================================================}
{�� _GroupFieldToItem(Table1 , 5 , combobox1.items  , 1);    }
procedure _GroupFieldToItem(TB:TTable ; FieldIndex:integer ; Item:Tstrings ; SpaceItem:integer);
var MyBookmark: TBookmark;
    i,j:longint;
    FieldName,DisplayText:string;
begin
  item.clear;
  if SpaceItem<>0 then item.add('');

  MyBookmark := TB.GetBookMark;
  TB.disablecontrols;

  j:=0;
  for i := 0 to (TB.fieldCount-1) do
      begin
        if TB.fields[i].visible=False then continue;
        j:=j+1;
        if j=FieldIndex then
           begin
             FieldName:=TB.fields[i].fieldname;
             break;
           end;
      end;

  TB.first;
  while not TB.eof do
     begin
       DisplayText:=TB.fieldbyname(FieldName).displayText;
       j:=0;
       for i := 0 to (item.Count-1) do
           begin
             if uppercase(_AllTrim(DisplayText))=uppercase(_AllTrim(item.strings[i])) then
                begin
                  j:=1;
                  break;
                end;
           end;
       if j=0 then item.add(DisplayText);
       TB.next;
     end;

  TB.GotoBookMark(MyBookMark);
  TB.FreeBookmark(MyBookmark);
  TB.enablecontrols;

end;
{==============================================================================}
{�M��table,�����]�M���ݩ�}
procedure _TableEmpty(TB: TTable);
begin
  TB.first;
  while not TB.eof do TB.delete;
end;
{==============================================================================}
{�h�椧 ���displayText�� (�H�r�I��)�p�ƾ�,}
function _Count(TB:TDataSet ; CountFields,Digits:string):string;
var sum: array[1..20] of Single ;
    MyBookmark: TBookmark;
    i,j,k:integer;
    tmp:string;
begin

  for i := 1 to 20 do sum[i]:=0;

  MyBookmark := TB.GetBookMark;
  TB.disablecontrols;
  TB.first;

  while not TB.eof do
     begin
       j:=0;
       for i := 0 to (TB.fieldCount-1) do
           begin
             if TB.fields[i].visible=False then continue;
             j:=j+1;

             for k:=1 to _SubStrNum(CountFields,',') do
                 begin
                   if _strtoint(_SegStr(CountFields,k))=0 then continue;
                   if _strtoint(_SegStr(CountFields,k))=j then
                      begin
                      if (_AllTrim(TB.fields[i].displayText)<>'') and
                         (_filternumstr(TB.fields[i].displayText)<>'') then
                         begin
                           sum[k]:=sum[k]+strtofloat(TB.fields[i].displayText);
                         end;
                      end;
                 end;
           end;
       TB.next;
     end;

  TB.GotoBookMark(MyBookMark);
  TB.FreeBookmark(MyBookmark);
  TB.enablecontrols;

  tmp:='';
  for k:=1 to _SubStrNum(CountFields,',') do
      begin
        tmp:=tmp+_str(sum[k],_strtoint(_SegStr(Digits,k)))+',';
      end;

  Result:=_cutright(tmp,1); {�h���̫ᤧ�r�I}

end;
{==============================================================================}
{���g����}
procedure _TableGroup(TB,TB2:TTable ; MainFieldID,CountFields,Digits:string);
var sum: array[1..20] of Single ;
    MyBookmark: TBookmark;
    i,j,k,l:integer;
    tmp,oldkey:string;
begin
  _TableEmpty(TB2);

  for i := 1 to 20 do sum[i]:=0;

  MyBookmark := TB.GetBookMark;
  TB.disablecontrols;
  TB.first;
  oldkey:='';
  while not TB.eof do
    begin
      j:=0;
      for i := 0 to (TB.fieldCount-1) do
          begin
            if TB.fields[i].visible=False then continue;
            j:=j+1;
            if oldkey='' then  oldkey:=TB.fields[i].displayText;
            if (j=_strtoint(MainFieldID)) and (oldkey<>TB.fields[i].displayText) then
               begin
                 TB2.append;
                 TB2.fields[0].asstring:=oldkey;
                 for l:= 1 to _SubStrNum(CountFields,',') do
                     TB2.fields[l].asstring:=floattostr(sum[l]);
                 for l := 1 to 20 do sum[l]:=0;
                 oldkey:=TB.fields[i].displayText;
               end;

            for k:=1 to _SubStrNum(CountFields,',') do
                 begin
                   if _strtoint(_SegStr(CountFields,k))=0 then continue;
                   if _strtoint(_SegStr(CountFields,k))=j then
                      begin
                      sum[k]:=sum[k]+_strtofloat(TB.fields[i].displayText);
                      end;
                 end;

          end;

      TB.next;
    end;

  TB.GotoBookMark(MyBookMark);
  TB.FreeBookmark(MyBookmark);
  TB.enablecontrols;

end;
{==============================================================================}
function _FieldType(FT:TFieldType):string;
var r:string;
begin
  r:='';
  if FT=ftUnknown     then r:='ftUnknown';
  if FT=ftstring      then r:='ftstring,Max 255';
  if FT=ftSmallint    then r:='ftSmallint,2,-32768 to 32767';
  if FT=ftinteger     then r:='ftinteger,4,-2147483648 to 2147483647';
  if FT=ftWord        then r:='ftWord,2,0 to 65535';
  if FT=ftBoolean     then r:='ftBoolean,2';
  if FT=ftFloat       then r:='ftFloat,8';
  if FT=ftCurrency    then r:='ftCurrency,8';
  if FT=ftBCD         then r:='ftBCD,18';
  if FT=ftDate        then r:='ftDate,4';
  if FT=ftTime        then r:='ftTime,4';
  if FT=ftDateTime    then r:='ftDateTime,8';
  if FT=ftBytes       then r:='ftBytes';
  if FT=ftVarBytes    then r:='ftVarBytes';
  if FT=ftAutoInc     then r:='ftAutoInc';
  if FT=ftBlob        then r:='ftBlob';
  if FT=ftMemo        then r:='ftMemo';
  if FT=ftGraphic     then r:='ftGraphic';
  if FT=ftFmtMemo     then r:='ftFmtMemo';
  if FT=ftParadoxOle  then r:='ftParadoxOle';
  if FT=ftDBaseOle    then r:='ftDBaseOle';
  if FT=ftTypedBinary then r:='ftTypedBinary';

  Result:=r;

end;
{==============================================================================}
function _TableNameExist(AliasName,TableName:string):Boolean;
var MyList: TstringList;
    r:Boolean;
begin
  r:=False;
  MyList := TstringList.Create;
  Session.GetTableNames(AliasName,'*.*', True, False, MyList);

  r:=_ItemExist(MyList,TableName);
  MyList.free;

  Result:=r;
end;
{==============================================================================}
{�ƻs���c,�L�k�ƻs����}
procedure _TableStructureCopy(Source,Dest:TTable ; TableName:string);
begin
  Dest.DatabaseName:=Source.DatabaseName;
  Dest.TableType:=Source.TableType;
  Dest.TableName:=TableName;

  if _AliasDefaultDriver(Dest.DatabaseName)='DBASE' then Dest.TableType:=ttDBase;
  if _AliasDefaultDriver(Dest.DatabaseName)='PARADOX' then Dest.TableType:=ttParadox;
  if _AliasDefaultDriver(Dest.DatabaseName)='ASCIIDRV' then Dest.TableType:=ttASCII;

  Dest.BatchMove(Source, batCopy);
  Dest.open;
  _tableempty(Dest);
  Dest.close;
end;
{==============================================================================}
{�P�W,�i�ƻs���c,�L�k�ƻs����}
procedure _TableNew(Source,Dest:TTable ; TableName:string);
var i:integer;
    name,fields,expression:string;
    options:TIndexOptions;
begin
  Dest.Active := False;
  Dest.DatabaseName:=Source.DatabaseName;
  Dest.TableName:=TableName;

  Dest.TableType:=Source.TableType;
  if _AliasDefaultDriver(Dest.DatabaseName)='DBASE' then Dest.TableType:=ttDBase;
  if _AliasDefaultDriver(Dest.DatabaseName)='PARADOX' then Dest.TableType:=ttParadox;
  if _AliasDefaultDriver(Dest.DatabaseName)='ASCIIDRV' then Dest.TableType:=ttASCII;

  Source.FieldDefs.update;
  Dest.FieldDefs:=Source.FieldDefs;
  Dest.CreateTable;

  {
  Source.IndexDefs.update;
  Dest.IndexDefs.clear;
  for i :=0 to (Source.IndexDefs.Count-1) do
    begin
      name:=Source.IndexDefs.items[i].name;
      fields:=Source.IndexDefs.items[i].fields;
      expression:=Source.IndexDefs.items[i].expression;
      if fields='' then fields:=expression;
      //showmessage('name : '+name+' fields : '+fields);
      options:=Source.IndexDefs.items[i].options;
      Dest.AddIndex(name,fields,options);
    end;   }

  //Dest.CreateTable;
end;
//==============================================================================
// ***** DataBase Alias ��Ʈw�O�W�Ѽ�
//==============================================================================
{�Ǧ^DATABASE ��m ----> ���|�θ�Ʈw}
function _AliasDataBase(AliasName:string):string; {�Ǧ^���|�̫�L'\'�r��}
var P:TstringList;
    Path:string;
begin
  P:=TstringList.Create;
  try
    begin
    Session.GetAliasParams(AliasName,P);

    Path:=P.Values['PATH'];
       if _right(Path,1)='\' then Path:=_cutright(Path,1);

    if _rtrim(Path)='' then
       begin
       Path:=P.Values['SERVER NAME'];
       //PARA:=_cutcharright(PARA,'\',1);
       end;
    end
  except
    begin
    showmessage(AliasName+'  �O�W���s�b');
    Path:='';
    end;
  end;

  P.free;
  Result:=Path;

end;
{==============================================================================}
{�Ǧ^ DATABASE ���|��m ----> �������| }
function _AliasDataBasePath(AliasName:string):string; {�Ǧ^���|�̫�L'\'�r��}
var P:TstringList;
    Path:string;
begin
  P:=TstringList.Create;
  try
    begin
    Session.GetAliasParams(AliasName,P);

    Path:=P.Values['PATH'];
       if _right(Path,1)='\' then Path:=_cutright(Path,1);

    if _rtrim(Path)='' then
       begin
       Path:=P.Values['SERVER NAME'];
       Path:=_cutcharright(Path,'\',1);
       end;
    end
  except
    begin
    showmessage(AliasName+'  �O�W���s�b');
    Path:='';
    end;
  end;

  P.free;
  Result:=Path;

end;
{==============================================================================}
function _Alias(AliasName:string):Boolean;
begin
  if Session.IsAlias(AliasName) then
     Result:=True
  else
     Result:=False;
end;
{==============================================================================}
function _AliasType(AliasName:string):string;
begin
  Result:=Session.GetAliasDriverName(AliasName);
end;
{==============================================================================}
function _AliasDefaultDriver(AliasName:string):string;
var MyList: TstringList;
    r:string;
begin
  r:='';
  MyList := TstringList.Create;
  Session.GetAliasParams(AliasName,MyList);
  r:=_stringListValue(MyList,'Default Driver');
  MyList.free;
  Result:=r;
end;
{==============================================================================}
procedure _AliasRemove(AliasName:string);
begin
  if _Alias(AliasName) then
     begin
     Session.DeleteAlias(AliasName);
     Session.SaveConfigFile; {�g�JBDE}
     end;
end;
{==============================================================================}
procedure _AliasCreate(AliasName,Path:string ; Driver:integer);
var DefaultDriver:string;
begin
  _AliasRemove(AliasName);

  if Path='' then Path:=extractfilepath(application.exename);
  if not Driver in [1..3] then Driver:=1;

  case Driver of
     1:DefaultDriver:='DBASE';
     2:DefaultDriver:='PARADOX';
     3:DefaultDriver:='ASCIIDRV';
  end;

  Session.AddStandardAlias(AliasName, Path, DefaultDriver);
  Session.SaveConfigFile; {�g�JBDE}
end;
{==============================================================================}
procedure _AliasSetup(AliasName,Path:string ; Driver:integer);
var DefaultDriver:string;
begin
  if _Alias(AliasName) then exit;

  _AliasCreate(AliasName,Path,Driver);
end;

function _AliasAppIni(AliasName:string):string;
var r,path:string;
begin
  r:='';
  path:=_AliasDataBasePath(AliasName);
  if _right(path,1)='\' then path:=_cutright(path,1);
  if path<>'' then r:=(path+'\'+application.title+'.INI');
  Result:=r;
end;
{==============================================================================}
{�H AlliasName �Ҧb���|(�p�����W) �� AppIni �� Valiable �`�Ϭ��ܼ��u��Ū�J��V}
{�Y AlliasName ���|���s�b(�p�����_�u) �h�H AP �Ҧb��m�� AppIni �� Valiable �`�Ϭ��ܼƳƥ�Ū�J��V}
{Example: showmessage(_ReadStr2('TICKET_R','TEST','1'));}
function _ReadAliasAppIniStr(AliasName,VarName,DefauVal:string):string;
var AppIni:string;
begin
  if _file(_AliasAppIni(AliasName)) then
     AppIni:=_AliasAppIni(AliasName)
  else
     AppIni:=_AppIni;

  Result:=_ReadWriteINIStr(AppIni,'Variable',VarName,DefauVal);
end;
{==============================================================================}
{�P�W,��Ū�J��ƭ�}
function _ReadAliasAppIniInt(AliasName,VarName:string ; DefauVal:integer):integer;
var AppIni:string;
begin
  if _file(_AliasAppIni(AliasName)) then
     AppIni:=_AliasAppIni(AliasName)
  else
     AppIni:=_AppIni;

  Result:=_strtoint(_ReadWriteINIStr(AppIni,'Variable',VarName,inttostr(DefauVal)));
end;
//==============================================================================
// ***** HTML �䴩���
//==============================================================================
{example : memo1.lines.add(_TableToHTMLContent(Table2,'Header','Footer',1));   }
{       or response.content:=_TableToHTMLContent(Table1,'Header','Footer',1);  }
function _TableToHTMLContent(TB:TDataSet ; header,footer:string ; fullHTML:integer):string;
var tablebody:string;
    tmp:TstringList;
    i,j:integer;
    LR:string;
    font_face:string;
begin
  tmp:=TstringList.Create;
  LR:=chr(13)+chr(10);
  font_face:='<font face="System">';

  tablebody:='';

  if fullHTML<>0 then
     begin
       tablebody:=tablebody+'<html>'+LR;

       tablebody:=tablebody+'<head>'+LR;
       tablebody:=tablebody+'<meta http-equiv="Content-Type" content="text/html; charset=big5">'+LR;
       tablebody:=tablebody+'<meta name="GENERATOR" content="Microsoft FrontPage Express 2.0">'+LR;
       tablebody:=tablebody+'<title>���R�W �@��e��</title>'+LR;
       tablebody:=tablebody+'</head>'+LR;

       tablebody:=tablebody+'<body bgcolor="#FFFFFF">'+LR+LR;
     end;

  {--------------------------------}

  if header<>'' then
  tablebody:=tablebody+'<p align="center"><font face="System">'+header+'</font></p>'+LR;

  tablebody:=tablebody+'<div align="center"><center>'+LR;
  tablebody:=tablebody+'<table border="1">'+LR;

  tablebody:=tablebody+'   <tr>'+LR;
    for i:=0 to TB.fieldCount-1 do
        begin
          tablebody:=tablebody+'      <td>';
          tablebody:=tablebody+font_face+TB.fields[i].fieldname+'</font>';
          tablebody:=tablebody+'</td>'+LR;
        end;
  tablebody:=tablebody+'   </tr>'+LR;

  TB.first;
  while not TB.eof do
        begin
          tablebody:=tablebody+'   <tr>'+LR;

          for i:=0 to TB.fieldCount-1 do
              begin
                tablebody:=tablebody+'      <td>';
                case TB.fields[i].datatype of
                  ftMemo : begin
                             tmp.assign(TB.fields[i]);

                             if tmp.Count<>0 then
                                begin
                                  for j:=0 to tmp.Count-1 do
                                      begin
                                        tablebody:=tablebody+font_face+tmp.strings[j]+'</font>';
                                        if i<>tmp.Count-1 then tablebody:=tablebody+'<br>'+LR+'          ';
                                      end;
                                end
                             else tablebody:=tablebody+font_face+'�@'+'</font>';//�ɥ����ť�

                           end;
                  {
                  ftstring : tablebody:=tablebody+font_face+TB.fields[i].DisplayText+'</font>';
                  ftinteger : ;
                  ftdate : ;
                  ftFloat : ;}
                else
                  begin
                    if _AllTrim(TB.fields[i].DisplayText)='' then
                       tablebody:=tablebody+font_face+'�@'+'</font>'//�ɥ����ť�
                    else
                       tablebody:=tablebody+font_face+TB.fields[i].DisplayText+'</font>';
                  end;
                end;
                tablebody:=tablebody+'</td>'+LR;
              end;

          tablebody:=tablebody+'   </tr>'+LR;
          TB.next;
        end;

  tablebody:=tablebody+'</table>'+LR;
  tablebody:=tablebody+'</center></div>'+LR;
  if footer<>'' then tablebody:=tablebody+'<p align="center"><font face="System">'+footer+'</font></p>'+LR;

  {--------------------------------}

  if fullHTML<>0 then
     begin
       tablebody:=tablebody+LR;
       tablebody:=tablebody+'</body>'+LR;
       tablebody:=tablebody+'</html>'+LR;
     end;

  tmp.free;
  Result:=tablebody;

end;
{==============================================================================}
{example : memo1.lines.add(_TableToHTMLEdit(Table2,'test','F',1,'MY_AP','3',2));}
{     or response.content:=_TableToHTMLEdit(Table2,'test','F',1,'MY_AP','3',2));}
{     �� : �N Table2 ���ĤT������,�é�̫�[�@������                           }
function _TableToHTMLEdit(TB:TDataSet ; header,footer:string ; fullHTML:integer ;
                          my_ap:string ; hide_field:string ; new_field:integer):string;
var tablebody:string;
    tmp:TstringList;
    i,j:integer;
    index,LR:string;
    font_face:string;
begin
  tmp:=TstringList.Create;
  LR:=chr(13)+chr(10);
  font_face:='<font face="System">';

  tablebody:='';

  if fullHTML<>0 then
     begin
       tablebody:=tablebody+'<html>'+LR;

       tablebody:=tablebody+'<head>'+LR;
       tablebody:=tablebody+'<meta http-equiv="Content-Type" content="text/html; charset=big5">'+LR;
       tablebody:=tablebody+'<meta name="GENERATOR" content="Microsoft FrontPage Express 2.0">'+LR;
       tablebody:=tablebody+'<title>���R�W �@��e��</title>'+LR;
       tablebody:=tablebody+'</head>'+LR;

       tablebody:=tablebody+'<body bgcolor="#FFFFFF">'+LR+LR;
     end;

  {--------------------------------}

  if header<>'' then
  tablebody:=tablebody+'<p align="center"><font face="System">'+header+'</font></p>'+LR+LR;

  tablebody:=tablebody+'<form action='+my_ap+' method="POST">'+LR;

  tablebody:=tablebody+'    <div align="center"><center>'+LR;
  tablebody:=tablebody+'    <table border="0">'+LR;

  if new_field=1 then  //�}�Y�s�[�@�������
     begin
       tablebody:=tablebody+'        <tr>'+LR;
       tablebody:=tablebody+'            <td>'+'<p align="right">'+font_face+'( TmpField )'+'</font></p></td>'+LR;
       tablebody:=tablebody+'            <td>'+font_face+'<input type="text" size="20"'+LR;
       tablebody:=tablebody+'                '+' name="V0"></font></td>'+LR;
       tablebody:=tablebody+'        </tr>'+LR;
     end;

  for i:=0 to TB.fieldCount-1 do
      begin
        index:=inttostr(i+1);

        if _strlike(index,hide_field) then continue;

        tablebody:=tablebody+'        <tr>'+LR;

        case TB.fields[i].datatype of
             ftMemo : begin
                        tmp.assign(TB.fields[i]);

                        tablebody:=tablebody+'            <td>'+'<p align="right">'+font_face+TB.Fields[i].DisplayLabel+'</font></p></td>'+LR;
                        tablebody:=tablebody+'            <td>'+font_face+'<textarea name="V'+index+'" rows="'+inttostr(_max(tmp.Count,2))+'"'+LR;
                        tablebody:=tablebody+'                '+' cols="40">'+LR;

                        if tmp.Count<>0 then
                           begin
                             for j:=0 to tmp.Count-1 do
                                 begin
                                   tablebody:=tablebody+tmp.strings[j];
                                   if j<>tmp.Count-1 then tablebody:=tablebody+chr(9)+LR;
                                 end;
                           end;

                        tablebody:=tablebody+'                </textarea></font></td>'+LR;

                      end;
        else
             begin
               tablebody:=tablebody+'            <td>'+'<p align="right">'+font_face+TB.Fields[i].DisplayLabel+'</font></p></td>'+LR;
               tablebody:=tablebody+'            <td>'+font_face+'<input type="text" size="'+inttostr(_max(TB.Fields[i].size,20))+'"'+LR;
               tablebody:=tablebody+'                '+' name="V'+index+'" value="'+TB.Fields[i].DisplayText+'"></font></td>'+LR;
             end;
        end;

        tablebody:=tablebody+'        </tr>'+LR;
      end;

  if new_field=2 then  //���ڷs�[�@�������
     begin
       index:=_serialNumAdd(index);
       tablebody:=tablebody+'        <tr>'+LR;
       tablebody:=tablebody+'            <td>'+'<p align="right">'+font_face+'( TmpField )'+'</font></p></td>'+LR;
       tablebody:=tablebody+'            <td>'+font_face+'<input type="text" size="20"'+LR;
       tablebody:=tablebody+'                '+' name="V'+index+'"></font></td>'+LR;
       tablebody:=tablebody+'        </tr>'+LR;
     end;

  tablebody:=tablebody+'        <tr>'+LR;
  tablebody:=tablebody+'             <td>�@</td>'+LR;
  tablebody:=tablebody+'             <td>�@</td>'+LR;
  tablebody:=tablebody+'        </tr>'+LR;

  tablebody:=tablebody+'        <tr>'+LR;
  tablebody:=tablebody+'             <td>�@</td>'+LR;
  tablebody:=tablebody+'             <td><input type="reset" name="B1" value="���]">�@'+LR;
  tablebody:=tablebody+'                 <input type="submit" name="B2" value="�ǰe"></td>'+LR;
  tablebody:=tablebody+'        </tr>'+LR;

  tablebody:=tablebody+'    </table>'+LR;
  tablebody:=tablebody+'    </center></div>'+LR;
  tablebody:=tablebody+'</form>'+LR;
  tablebody:=tablebody+LR;

  if footer<>'' then tablebody:=tablebody+'<p align="center"><font face="System">'+footer+'</font></p>'+LR;

  {--------------------------------}

  if fullHTML<>0 then
     begin
       tablebody:=tablebody+LR;
       tablebody:=tablebody+'</body>'+LR;
       tablebody:=tablebody+'</html>'+LR;
     end;

  tmp.free;
  Result:=tablebody;

end;
{==============================================================================}
function _StrGridToHTML(SG:TstringGrid ; header,footer:string ; fullHTML:integer ;
                        my_ap:string ; TableWidth:string ; hide_field:string ; key_field:integer ):string ;
var tablebody:string;
    tmp:TstringList;
    i,j,k:integer;
    TW,LR,CELL,subCELL:string;
    font_face:string;
    kl:integer;
begin
  tmp:=TstringList.Create;
  LR:=chr(13)+chr(10);
  font_face:='<font face="System">';

  tablebody:='';

  if fullHTML<>0 then
     begin
       tablebody:=tablebody+'<html>'+LR;

       tablebody:=tablebody+'<head>'+LR;
       //tablebody:=tablebody+'<meta http-equiv="Content-Type" content="text/html; charset=big5">'+LR;
       //tablebody:=tablebody+'<meta name="GENERATOR" content="Microsoft FrontPage Express 2.0">'+LR;
       tablebody:=tablebody+'<title>���R�W �@��e��</title>'+LR;
       tablebody:=tablebody+'</head>'+LR;

       tablebody:=tablebody+'<body bgcolor="#FFFFFF">'+LR+LR;
     end;

  if key_field<>0 then
     begin
       my_ap:=extractfilename(my_ap);
       tablebody:=tablebody+'<form action="'+my_ap+'" method="GET" enctype="text/plain">'+LR+LR;
     end;

  {--------------------------------}

  if header<>'' then
  tablebody:=tablebody+'<p align="center"><font face="System">'+header+'</font></p>'+LR;

  tablebody:=tablebody+'<div align="center"><center>'+LR+LR;  {TABLE �襤}

  TW:=TableWidth; if TW='' then TW:='100%';

  tablebody:=tablebody+'<table border="1" width="'+TW+'">'+LR;

  {
  if key_field<>0 then    //�Ʊ�C�ӫ���e�שT�w�����[,���դF�L��
     begin
       kl:=0;
       for j:=0 to SG.rowCount-1 do
           begin
             if Length(SG.cells[key_field-1,j])>kl then kl:=Length(SG.cells[key_field-1,j]);
           end;
     end; }

  for j:=0 to SG.rowCount-1 do
      begin

        tablebody:=tablebody+'   <tr>'+LR;

        for i:=0 to SG.colCount-1 do
            begin

              if _strlike(inttostr(i+1),hide_field) then continue;

              tablebody:=tablebody+'      <td>';

              CELL:=SG.cells[i,j];if CELL='' then CELL:='�@';//�ɥ����ť�


              if ((i+1)=key_field) and ((j+1)>sg.fixedrows) and (CELL<>'�@') and (CELL<>' ') then
                 begin
                   tablebody:=tablebody+'<input type="submit" name="'+'key'{SG.cells[i,0]}+'" value="'+CELL+'">';
                 end
              else
                 begin
                   if _SubStrNum(CELL,'||')<>0 then
                      begin
                        for k:=1 to (_SubStrNum(CELL,'||')+1) do
                            begin
                              subCELL:=_GetSegStr(CELL,'||',k);
                              tablebody:=tablebody+font_face+_HypLnkText(subCELL,'')+'</font>';
                              if K<>(_SubStrNum(CELL,'||')+1) then tablebody:=tablebody+'<br>'+LR+'          ';
                            end
                      end
                   else tablebody:=tablebody+font_face+_HypLnkText(CELL,'')+'</font>';
                 end;


              tablebody:=tablebody+'</td>'+LR;
            end;

        tablebody:=tablebody+'   </tr>'+LR;
      end;

  tablebody:=tablebody+'</table>'+LR+LR;

  tablebody:=tablebody+'</center></div>'+LR; {TABLE �襤}

  if footer<>'' then tablebody:=tablebody+'<p align="center"><font face="System">'+footer+'</font></p>'+LR;

  {--------------------------------}

  if key_field<>0 then
     begin
       tablebody:=tablebody+LR+'</form>'+LR;
     end;

  if fullHTML<>0 then
     begin
       tablebody:=tablebody+LR;
       tablebody:=tablebody+'</body>'+LR;
       tablebody:=tablebody+'</html>'+LR;
     end;

  tmp.free;
  Result:=tablebody;

end;
{==============================================================================}
function _StrGridToHTMLEdit(SG:TstringGrid ; header,footer:string ; fullHTML:integer ;
                            my_ap:string ; hide_key , hide_field , readonly_field ,
                            new_headfield , new_footfield:string):string;
var tablebody,cell:string;
    tmp:TstringList;
    i,j:integer;
    index,LR:string;
    font_face:string;
begin
  tmp:=TstringList.Create;
  LR:=chr(13)+chr(10);
  font_face:='<font face="System">';

  tablebody:='';

  if fullHTML<>0 then
     begin
       tablebody:=tablebody+'<html>'+LR;

       tablebody:=tablebody+'<head>'+LR;
       //tablebody:=tablebody+'<meta http-equiv="Content-Type" content="text/html; charset=big5">'+LR;
       //tablebody:=tablebody+'<meta name="GENERATOR" content="Microsoft FrontPage Express 2.0">'+LR;
       tablebody:=tablebody+'<title>���R�W �@��e��</title>'+LR;
       tablebody:=tablebody+'</head>'+LR;

       tablebody:=tablebody+'<body bgcolor="#FFFFFF">'+LR+LR;
     end;

  {--------------------------------}

  if header<>'' then
     tablebody:=tablebody+'<p align="center"><font face="System">'+header+'</font></p>'+LR+LR;

  tablebody:=tablebody+'<form action='+extractfilename(my_ap)+'/EDIT'+' method="GET" enctype="text/plain">'+LR;

  if hide_key<>'' then //�]�w���� KEY
     tablebody:=tablebody+'<input type="hidden" name="EditKey" value="'+hide_key+'">'+LR;

  tablebody:=tablebody+'    <div align="center"><center>'+LR;
  tablebody:=tablebody+'    <table border="0">'+LR;

  if new_headfield<>'' then  //�}�Y�s�[�@�������
     begin
       tablebody:=tablebody+'        <tr>'+LR;
       tablebody:=tablebody+'            <td>'+'<p align="right">'+font_face+new_headfield+'</font></p></td>'+LR;
       tablebody:=tablebody+'            <td>'+font_face+'<input type="text" size="20"'+LR;
       tablebody:=tablebody+'                '+' name="V0"></font></td>'+LR;
       tablebody:=tablebody+'        </tr>'+LR;
     end;

  for i:=0 to SG.colCount-1 do
      begin
        index:=inttostr(i+1);
        cell:=SG.cells[i,SG.row];

        if _strlike(index,hide_field) then continue;

        tablebody:=tablebody+'        <tr>'+LR;

        if _strlike(index,readonly_field) then
           begin
             tablebody:=tablebody+'            <td>'+'<p align="right">'+font_face+SG.cells[i,0]+'</font></p></td>'+LR;
             tablebody:=tablebody+'            <td>'+font_face+cell+'</font></td>'+LR;
           end
        else
           begin
             tablebody:=tablebody+'            <td>'+'<p align="right">'+font_face+SG.cells[i,0]+'</font></p></td>'+LR;
             tablebody:=tablebody+'            <td>'+font_face+'<input type="text" size="'+inttostr(_max(Length(cell),20))+'"'+LR;
             tablebody:=tablebody+'                '+' name="V'+index+'" value="'+cell+'"></font></td>'+LR;
           end;

        tablebody:=tablebody+'        </tr>'+LR;
      end;

  if new_footfield<>'' then  //���ڷs�[�@�������
     begin
       index:=_serialNumAdd(index);
       tablebody:=tablebody+'        <tr>'+LR;
       tablebody:=tablebody+'            <td>'+'<p align="right">'+font_face+new_footfield+'</font></p></td>'+LR;
       tablebody:=tablebody+'            <td>'+font_face+'<input type="text" size="20"'+LR;
       tablebody:=tablebody+'                '+' name="V'+index+'"></font></td>'+LR;
       tablebody:=tablebody+'        </tr>'+LR;
     end;

  tablebody:=tablebody+'        <tr>'+LR;
  tablebody:=tablebody+'             <td>�@</td>'+LR;
  tablebody:=tablebody+'             <td>�@</td>'+LR;
  tablebody:=tablebody+'        </tr>'+LR;

  tablebody:=tablebody+'        <tr>'+LR;
  tablebody:=tablebody+'             <td>�@</td>'+LR;
  tablebody:=tablebody+'             <td><input type="reset" name="B1" value="���]">�@'+LR;
  tablebody:=tablebody+'                 <input type="submit" name="B2" value="�ǰe"></td>'+LR;
  tablebody:=tablebody+'        </tr>'+LR;

  tablebody:=tablebody+'    </table>'+LR;
  tablebody:=tablebody+'    </center></div>'+LR;
  tablebody:=tablebody+'</form>'+LR;
  tablebody:=tablebody+LR;

  if footer<>'' then tablebody:=tablebody+'<p align="center"><font face="System">'+footer+'</font></p>'+LR;

  {--------------------------------}

  if fullHTML<>0 then
     begin
       tablebody:=tablebody+LR;
       tablebody:=tablebody+'</body>'+LR;
       tablebody:=tablebody+'</html>'+LR;
     end;

  tmp.free;
  Result:=tablebody;

end;
{==============================================================================}
function _HypLnkText(C,DispText:string):string;
var i:integer;
    r:string;
    check:Boolean;
    Disp:string;
begin

  if DispText='' then Disp:=c
     else Disp:=DispText;

  check:=False;
  for i:=1 to Length(C) do   // �����r�X�~�P�� e_mail �r�X
      begin
        //if C[i]>chr(127) then
        if ord(copy(C,i,1)[1])>127 then
           begin
             check:=True;
             Break;
           end;
      end;
  if check then
     begin
       Result:=Disp;
       Exit;
     end;


  if _IsEmail(C) then
     r:='<a href="mailto:'+C+'">'+Disp+'</a>'
  else
    begin
      if UpperCase(_Left(C,7))='HTTP://' then
         r:='<a href="http://'+_CutLeft(C,7)+'">'+Disp+'</a>'
      else
         r:=Disp;
    end;

  Result:=r;

end;

{==============================================================================}
function _HTMLConfirmContent(msg,http:string):string;
var LR,Content:string;
begin

  LR:=chr(13)+chr(10);

  Content:='';
  {
  Content:=Content+'<html>'+LR;
  Content:=Content+'<head>'+LR;
  Content:=Content+'  <title>���R�W �@��e��</title>'+LR;
  Content:=Content+'</head>'+LR;
  Content:=Content+'<body bgcolor="#FFFFFF">'+LR;

  Content:=Content+'<form action="'+http+'" method="POST">'+LR;
  Content:=Content+'  <p align="center">�@</p>'+LR;
  Content:=Content+'  <p align="center">�@</p>'+LR;
  Content:=Content+'  <p align="center">'+msg+'</p>'+LR;
  Content:=Content+'  <p align="center">�@</p>'+LR;
  Content:=Content+'  <p align="center">�@</p>'+LR;
  Content:=Content+'  <p align="center"><input type="submit" name="B1" value="�@OK�@"></p>'+LR;
  Content:=Content+'</form>'+LR;

  Content:=Content+'</body>'+LR;
  Content:=Content+'</html>'+LR;
   }

  Content:=Content+'<html>'+LR;
  Content:=Content+'<head>'+LR;
  Content:=Content+'<title>���R�W �@��e��</title>'+LR;
  Content:=Content+'</head>'+LR;
  Content:=Content+'<body bgcolor="#FFFFFF">'+LR;

  Content:=Content+'<p>�@</p>'+LR;
  Content:=Content+'<p align="center"><font face="System">'+msg+'</font></p>'+LR;
  Content:=Content+'<p align="center"><font face="System"></font>�@</p>'+LR;
  Content:=Content+'<p align="center"><font face="System"></font>�@</p>'+LR;
  Content:=Content+'<p align="center"><font face="System"></font>�@</p>'+LR;
  Content:=Content+'<p align="center"><a href="'+http+'"><font face="System">�T�w</font></a></p>'+LR;

  Content:=Content+'</body>'+LR;
  Content:=Content+'</html>'+LR;

  Result:=Content;

end;
{==============================================================================}
function _strToASCII(s:string):string;
var ascii:string;
    i:integer;
begin
  ascii:='';
  for i:=1 to Length(s) do
      begin
        ascii:=ascii+inttostr(ord(s[i]))+' ';
      end;
      
  Result:=ascii;
end;
{==============================================================================}
function _BigSpace(s:string):string; {�Y���Ŧr��h�Ǧ^�����ť�}
begin
  if s='' then
    Result:='�@'
  else
    Result:=s;
end;
{==============================================================================}
function _CounterToHtml:string;  {���Y�@HTML Table cell ��}
var Counter,TodayCounter:LongInt;
    CountFrom,TodayCounterFrom:string;
    LR:string;
    Hour:integer;
    TodayTimeHit:string;
begin

  LR:=chr(13)+chr(10);

  Counter:=_QRInt('Counter',0);
  if Counter=0 then _QWStr('CountFrom',_date(0,'.')); //�j��g�J
  Counter:=Counter+1 ; _QWInt('Counter',Counter);
  CountFrom:=_QRStr('CountFrom','');

  TodayCounter:=_QRInt('TodayCounter',0); {����h�l,���F���� INI �ɤ� keyword ����}
  TodayCounterFrom:=_QRStr('TodayCountFrom','');
  if TodayCounterFrom<>_date(0,'.') then
     begin
       _QWInt('TodayCounter',0);
       _QWStr('TodayCountFrom',_date(0,'.'));
       _QWStr('TodayTimeHit','');
     end;
  TodayCounter:=_QRInt('TodayCounter',0);
  TodayCounter:=TodayCounter+1 ; _QWInt('TodayCounter',TodayCounter);

  TodayTimeHit:=_QRStr('TodayTimeHit','');
  Hour:=_StrToInt(_WhatTime('Hour'));
  _SegStrValueInc(TodayTimeHit,',',Hour+1); {00:??:?? �p�ɦb�Ĥ@�`}
  _QWStr('TodayTimeHit',TodayTimeHit);

  Result:='<font face="System">�����s���G</font><font color="#FF0000" face="System">'+_strzero(TodayCounter,6)+'</font><font face="System"> �H</font><br>'+
          '<font face="System">�֭p�s���G</font><font color="#FF0000" face="System">'+_strzero(Counter,6)+'</font><font face="System"> �H</font><br>'+
          '<font face="System">�_�����G'+CountFrom+'</font>'

end;
{==============================================================================}
function _HTMLTimeHit(TimeHitStr:string):string;
var LR:string;
    i,j,v,max:integer;
    H:array[1..24 , 1..200] of string;
    t,r:string;
begin

  if _AllTrim(TimeHitStr)='' then TimeHitStr:='TimeHit'; { ���w��_TimeHit() ��Ʀs�ɸ��}

  LR:=chr(13)+chr(10);
  max:=0;
  for i:=1 to 24 do
      begin
        v:=_strtoint(_GetSegStr(TimeHitStr,',',i));
        if v>max then max:=v;
      end;

  //r:='<p>�@</p>'+LR;
  r:=r+'<p align="center">'+LR;
  r:=r+'<font color="#0000FF" font face="system">'+LR;
  for j := max+1 downto 1 do
      begin
        t:='';
        for i := 1 to 24 do
            begin
              v:=_strtoint(_GetSegStr(TimeHitStr,',',i));

              if j=v+1 then
                 begin
                   if v<>0 then t:=t+_strzero(v,2)+' ' else t:=t+'�� ';
                 end
              else
                 begin
                   if j<=v then t:=t+'�i ' else t:=t+'�� ';
                 end;
            end;
         r:=r+'   '+t+'<br>'+LR;
      end;
   r:=r+'</font>'+LR;
   r:=r+'<font face="system">00 01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19 20 21 22 23'+'<br>'+LR;
   r:=r+'</font>'+LR;
   r:=r+'</p>'+LR+LR;

   r:=r+'<p align="center"><font color="#FF0000" face="system">'+'����]�P��'+_WeekStr(date)+'�^�����s���H�� 24 �p�ɮɬq���G��'+'</font></p>'+LR+LR;

   Result:=r;

end;
{==============================================================================}
function _HTMLTimeHit2(TimeHitStr:string):string; //�A�˹�
var LR:string;
    i,j,v,max:integer;
    H:array[1..24 , 1..200] of string;
    t,r:string;
begin

  if _AllTrim(TimeHitStr)='' then TimeHitStr:='TimeHit'; { ���w��_TimeHit() ��Ʀs�ɸ��}

  LR:=chr(13)+chr(10);
  max:=0;
  for i:=1 to 24 do
      begin
        v:=_strtoint(_GetSegStr(TimeHitStr,',',i));
        if v>max then max:=v;
      end;

  r:=r+'<p align="center"><font color="#FF0000" face="system">'+'����]�P��'+_WeekStr(date)+'�^�����s���H�� 24 �p�ɮɬq���G��'+'</font></p>'+LR+LR;
  r:=r+'<p align="center"><font face="system">00 01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19 20 21 22 23'+'<br></font>'+LR;
  r:=r+'<font color="#0000FF" font face="system">'+LR;

  for j := 1 to max+1 do
      begin
        t:='';
        for i := 1 to 24 do
            begin
              v:=_strtoint(_GetSegStr(TimeHitStr,',',i));

              if j=v+1 then
                 begin
                   if v<>0 then t:=t+_strzero(v,2)+' ' else t:=t+'�� ';
                 end
              else
                 begin
                   if j<=v then t:=t+'�i ' else t:=t+'�� ';
                 end;
            end;
         r:=r+'   '+t+'<br>'+LR;
      end;
   r:=r+'</font>'+LR;

   r:=r+'</p>'+LR+LR;



   Result:=r;

end;
{==============================================================================}
procedure _DelTree(dir:String);  {2001/01/31}
var
FileRec:TSearchrec;
begin
  if Dir[Length(Dir)]<>'\' then Dir := Dir + '\';
  if FindFirst(Dir+'*.*',faAnyfile,FileRec) = 0 then
    repeat
      if (FileRec.Attr = faDirectory) then
         begin
           if (FileRec.Name<>'.') and (FileRec.Name<>'..') then
              begin
                DelTree(Dir+FileRec.Name);
                RemoveDir(Dir+FileRec.Name);
              end;
         end
      else
         begin
           FileSetAttr(Dir+FileRec.Name,faArchive);
           deletefile(Dir+FileRec.Name);
         end;
    until FindNext(FileRec)<>0;

  FindClose(FileRec);

end;
{==============================================================================}
{decimal ���w�ध�Q�i��� , nToCarry ���i��� (���o�j�� 16)}
function _ToCarryAs(decimal,nToCarry:longint):string;
var r:string;
surplus:longint;
begin
  surplus:=decimal;
  r:='';
  while surplus��=nToCarry do
        begin
          r:=inttohex((surplus mod nToCarry),0)+r;
          surplus:=round(int(surplus/nToCarry));
        end;
  r:=inttohex(surplus,0)+r;
  result:=r;
end;

{##############################################################################}
{ �Ѧҽg }
{##############################################################################}

{�U�����Ѩ�Ө禡�ѧֳt�g�JŪ�X ini �����ܼ�}
{function TForm1._ReadString(VarName:string):string;
var MyIni: TIniFile;
IniFileName:string;
begin
IniFileName:=ChangeFileExt(ParamStr(0),'.INI'); //���w���� ini �ɦW
MyIni := TIniFile.Create(IniFileName);
Result:=MyIni.ReadString('Parameter',VarName,'');
MyIni.Free;
end;

procedure TForm1._WriteString(VarName,VarValue:string);
var MyIni: TIniFile;
IniFileName:string;
begin
IniFileName:=ChangeFileExt(ParamStr(0),'.INI'); //���w���� ini �ɦW
MyIni := TIniFile.Create(IniFileName);
MyIni.WriteString('Parameter',VarName,VarValue);
MyIni.Free;

end;}



end.
