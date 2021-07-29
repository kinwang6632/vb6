unit dtmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, ADODB, DBTables, DBClient, IniFiles;

const
    INVALID_GROUP_ID = 999;
    APPINFO_FILENAME = 'AppInfo.txt';
type
  TGroupInfoObj = class(TObject)
    Group : String;
    TotalOccurrence : Integer;
    TotalPercent : double;
    MemberCount  :  Integer;
    ActivityCode  :  String;
    nOccurrence1: Integer;
    nOccurrence2: Integer;
    nOccurrence3: Integer;
  end;

  PAppInfo=^TAppInfo;
  TAppInfo = record
    AppId: String;
    AppName: String;
    SubAppId: String; 
  end;

  TdtmMain = class(TDataModule)
    qryActivityCodeStat: TQuery;
    qryActivityCodeStatActivityCode: TStringField;
    qryActivityCodeStatOccurrence: TIntegerField;
    cdsActivityCodeStat: TClientDataSet;
    cdsActivityCodeStatnOccurrence: TIntegerField;
    cdsActivityCodeStatsActivityCode: TStringField;
    cdsActivityCodeStatnGroup: TIntegerField;
    qryAgentPerformance: TQuery;
    qryAgentBySkillsetPerformance: TQuery;
    qryAgentBySkillsetPerformanceTalkTime: TIntegerField;
    qryAgentBySkillsetPerformanceNotReadyTime: TIntegerField;
    qryAgentBySkillsetPerformanceCallsAnswered: TIntegerField;
    qryAgentBySkillsetPerformanceWorkTime: TIntegerField;
    qryAgentBySkillsetPerformanceAgentSurName: TStringField;
    cdsAgentPerformance: TClientDataSet;
    cdsAgentPerformanceGroup: TIntegerField;
    cdsAgentPerformanceAgentLogin: TStringField;
    cdsAgentPerformanceCallsAnswered: TIntegerField;
    cdsAgentPerformanceTalkTime: TIntegerField;
    cdsAgentPerformanceNotReadyTime: TIntegerField;
    cdsAgentPerformanceAgentGivenName: TStringField;
    Table1: TTable;
    cdsActivityCodeStatsActivityCodeName: TStringField;
    qryActivityCodeStatActivityCodeName: TStringField;
    cdsAgentPerformanceCallsPresented: TIntegerField;
    cdsAgentPerformanceLoggedInTime: TIntegerField;
    Database1: TDatabase;
    cdsActivityCodeStatnOccurrence1: TIntegerField;
    cdsActivityCodeStatnOccurrence2: TIntegerField;
    cdsActivityCodeStatnOccurrence3: TIntegerField;
    cdsAgentPerformanceShortCallsAns: TIntegerField;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
    procedure Database1Login(Database: TDatabase; LoginParams: TStrings);
  private
    { Private declarations }
    L_ValidNtyCodeStrList : TStringList;
    L_ValidBestCodeStrList : TStringList;
    L_ValidShCodeStrList : TStringList;
    L_ValidCyCodeStrList : TStringList;
    L_GroupInfo : TStringList;
    procedure computeGroupInfo(nI_CompNdx,   nI_TotalOccurrence:Integer);
    procedure loadTmpCode;
    function getGroup(nI_CompNdx, nI_Code : Integer):Integer;
    procedure ReleaseAppInfoObject;
  public
    { Public declarations }
    procedure getGroupInfo(sI_Group:String; var nI_Occurrence,  nI_MemberCount:Integer; var fI_Percent:double;  var  sI_GroupActivityCode  :  String);
    procedure getGroupInfo2(sI_Group:String; var nIOccurrence1,
      nIOccurrence2, nIOccurrence3: Integer);
    function checkIfInvalid(nI_CompNdx : Integer;sI_ActivityCode:String):boolean;
    procedure initCDSItems(nI_CompNdx:Integer);
    function activeAgentBySkillsetPerformance(nI_CompNdx : Integer;dI_SDate, dI_EDate : TDate):boolean;    
    function activeAgentPerformanceStat(nI_DataSrc ,nI_CompNdx : Integer;dI_SDate, dI_EDate : TDate):boolean;    
    function activeActivityCodeStat(nI_DataSrc , nI_CompNdx : Integer;dI_SDate, dI_EDate : TDate; var nI_NtyTOccurrences, nI_BestTOccurrences,nI_ShTOccurrences,nI_CyTOccurrences: Integer):boolean;
//    function activeActivityCodeStat(nI_CompNdx : Integer;dI_SDate : TDate; var nI_NtyTOccurrences, nI_LscTOccurrences : Integer):boolean;
    procedure SaveAppObjToFile;
  end;

var
  dtmMain: TdtmMain;

  AppInfoList: TList;

implementation

{$R *.DFM}

function TdtmMain.activeActivityCodeStat(nI_DataSrc , nI_CompNdx : Integer;dI_SDate, dI_EDate: TDate; var nI_NtyTOccurrences, nI_BestTOccurrences,nI_ShTOccurrences,nI_CyTOccurrences : Integer): boolean;
//function TdtmMain.activeActivityCodeStat(nI_CompNdx : Integer;dI_SDate: TDate; var nI_NtyTOccurrences, nI_LscTOccurrences : Integer): boolean;
var
    sL_CompCond, sL_CodeName, sL_Code : String;
    nL_Group, nL_Code , nL_Occurrence, nL_TotalOccurrence : Integer;
    bL_HasRecord : boolean;
    wL_Year, wL_Month, wL_Day : Word;
    sL_SDate, sL_EDate, sL_YMD : String;
    sL_TableName : String;
    sL_ApplicationID  :  String;

    aAppCount: Integer;
    aIndex: Integer;

    function GetSubCount(aSub: String): Integer;
    var
      aIdx: Integer;
    begin
      Result := 0;
      if Length( aSub ) <= 0 then Exit;
      for aIdx := 1 to Length( aSub ) do
        if aSub[aIdx] = ',' then Inc( Result );
      Inc( Result );
    end;

    function GetSubAppIdByIndex(aSub: String; aAssignIdx: Integer): String;
    var
      aStrList: TStringList;
    begin
      Result := EmptyStr;
      aStrList := TStringList.Create;
      try
        aStrList.DelimitedText := ',';
        aStrList.CommaText := aSub;
        if ( aAssignIdx <= aStrList.Count ) then
          Result := aStrList[aAssignIdx-1];
      finally
        aStrList.Free;
      end;
    end;

begin
    decodeDate(dI_SDate,wL_Year, wL_Month, wL_Day);
    sL_YMD := Format('%.4d%.2d%.2d', [ wL_Year, wL_Month, wL_Day ]);
    sL_SDate := sL_YMD;

    case nI_DataSrc of
      0: //從Symposium的日報抓data
       begin
          decodeDate(dI_EDate,wL_Year, wL_Month, wL_Day); //0604
        sL_TableName := ' dActivityCodeStat ';
       end;
      1: //從Symposium的月報抓data
       begin
        decodeDate(dI_EDate,wL_Year, wL_Month, wL_Day); //因為Symposium的monthly report的date都是該月份的第一天,所以這麼寫程式(用dI_SDate來替代dI_EDate)
        sL_TableName := ' mActivityCodeStat ';
       end;
      2: //從Symposium的 interval report抓data
       begin
        decodeDate(dI_EDate,wL_Year, wL_Month, wL_Day);
        sL_TableName := ' iActivityCodeStat ';
       end;
    end;
    sL_YMD := Format('%.4d%.2d%.2d', [ wL_Year, wL_Month, wL_Day ]);
    sL_EDate := sL_YMD;


    sL_CompCond := ' ';

    {

    {
    Application  ID
    NTY                 10004
    BEST                10015
    SH                  10003
    CY                  10011
    LSC                 10007
    BEST_LSC            10012
    SH_LSC              10000
    CY_LSC              10008

    SERVICE_NTY         10005
    SERVICE_BEST        10014
    SERVICE_SH          10002
    SERVICE_CY          10010
    }

    { 新的
    NTY + LSC + SERVICE_NTY           10004 + 10007 + 10005
    BEST + BEST_LSC + SERVICE_BEST    10015 + 10012 + 10014
    SH + SH_LSC + SERVICE_SH          10003 + 10000 + 10002
    CY + CY_LSC + SERVICE_CY          10011 + 10008 + 10010

    LSC                 10007
    BEST_LSC            10012
    SH_LSC              10000
    CY_LSC              10008

    SERVICE_NTY         10005
    SERVICE_BEST        10014
    SERVICE_SH          10002
    SERVICE_CY          10010

    }

    aAppCount := GetSubCount( PAppInfo( AppInfoList[nI_CompNdx] ).SubAppId );

    bL_HasRecord := False;
    nL_TotalOccurrence := 0;

    loadTmpCode;

for aIndex := 1 to aAppCount do
begin
    {
    case nI_CompNdx of
      0: //NTY
        sL_ApplicationID  :=  '10004';
      1: //BEST
        sL_ApplicationID  :=  '10015';
      2: //SH
        sL_ApplicationID  :=  '10003';
      3: //CY
        sL_ApplicationID  :=  '10011';

    end;
    }
    {1023, 10:05
    case nI_CompNdx of
      0: //NTY
        sL_CompCond := ' and ActivityCode between "001" and "078"';
      1: //LSC
        sL_CompCond := ' and ActivityCode between "501" and "561"';
    end;
    }

    sL_ApplicationID := GetSubAppIdByIndex(
      PAppInfo( AppInfoList[nI_CompNdx] ).SubAppId, aIndex );


    with qryActivityCodeStat do
    begin
      SQL.Clear;
      SQL.Add('select distinct ActivityCode, ActivityCodeName, sum(Occurrences) Occurrence');
      SQL.Add('from ' + sL_TableName);
      SQL.Add('where convert( varchar(8), Timestamp, 112 ) between :StartDate and :EndDate');
      SQL.Add(sL_CompCond);
      SQL.Add(' and ApplicationID=' + sL_ApplicationID);
      SQL.Add('  group by ActivityCode');
      SQL.Add('order by ActivityCode');
      ParamByName('StartDate').AsString := sL_SDate;
      ParamByName('EndDate').AsString := sL_EDate;
      
      Open;
      if (RecordCount>0) then
      begin
        bL_HasRecord := True;
        if aIndex = 1 then
        begin
          initCDSItems(nI_CompNdx);
          cdsActivityCodeStat.First;
        end;

        while not Eof do
        begin
          nL_Code := FieldByName('ActivityCode').AsInteger;
          sL_Code := FieldByName('ActivityCode').AsString;
          sL_CodeName := FieldByName('ActivityCodeName').AsString;
          nL_Occurrence := FieldByName('Occurrence').AsInteger;
          if (cdsActivityCodeStat.Locate('sActivityCode',FieldByName('ActivityCode').AsString,[])) then
          begin
            cdsActivityCodeStat.Edit;
            cdsActivityCodeStat.FieldByName('nOccurrence').AsInteger :=
              ( cdsActivityCodeStat.FieldByName('nOccurrence').AsInteger + nL_Occurrence );
            case aIndex of
              1: cdsActivityCodeStat.FieldByName('nOccurrence1').AsInteger := nL_Occurrence;
              2: cdsActivityCodeStat.FieldByName('nOccurrence2').AsInteger := nL_Occurrence;
              3: cdsActivityCodeStat.FieldByName('nOccurrence3').AsInteger := nL_Occurrence;
            end;
            cdsActivityCodeStat.Post;
            nL_TotalOccurrence   :=   nL_TotalOccurrence +   nL_Occurrence;
            case nI_CompNdx of
              0: //NTY
                nI_NtyTOccurrences := nI_NtyTOccurrences + nL_Occurrence;
              1: //BEST
                nI_BestTOccurrences := nI_BestTOccurrences + nL_Occurrence;
              2: //SH
                nI_ShTOccurrences := nI_ShTOccurrences + nL_Occurrence;
              3: //CY
                nI_CyTOccurrences := nI_CyTOccurrences + nL_Occurrence;
            end;
          end
          else
          begin
              cdsActivityCodeStat.Append;
              cdsActivityCodeStat.FieldByName('nGroup').AsInteger := INVALID_GROUP_ID;
              cdsActivityCodeStat.FieldByName('sActivityCode').AsString := sL_Code;
              cdsActivityCodeStat.FieldByName('sActivityCodeName').AsString := sL_CodeName;
              cdsActivityCodeStat.FieldByName('nOccurrence').AsInteger := nL_Occurrence;
              cdsActivityCodeStat.FieldByName('nOccurrence1').AsInteger := 0;
              cdsActivityCodeStat.FieldByName('nOccurrence2').AsInteger := 0;
              cdsActivityCodeStat.FieldByName('nOccurrence3').AsInteger := 0;
            case aIndex of
              1: cdsActivityCodeStat.FieldByName('nOccurrence1').AsInteger := nL_Occurrence;
              2: cdsActivityCodeStat.FieldByName('nOccurrence2').AsInteger := nL_Occurrence;
              3: cdsActivityCodeStat.FieldByName('nOccurrence3').AsInteger := nL_Occurrence;
            end;
            cdsActivityCodeStat.Post;
          end;
          Next;
        end;
        Close;
        //computeGroupInfo(nI_CompNdx, nL_TotalOccurrence);
      end;
    end;
end;

computeGroupInfo(nI_CompNdx, nL_TotalOccurrence);
Result := bL_HasRecord;
end;

procedure TdtmMain.DataModuleCreate(Sender: TObject);
begin
    Database1.Connected := false;
    try
      Database1.Connected := true;
    except
      on E: Exception do
      MessageDlg( E.Message, mtError, [mbOK], 0 );
    end;
{
    Database1.Params.Clear;
    Database1.Params.Add('USERNAME=sysadmin');
    Database1.Params.Add('PASSWORD=nortel');
}

    L_ValidNtyCodeStrList := TStringList.Create;
    L_ValidBestCodeStrList := TStringList.Create;
    L_ValidShCodeStrList := TStringList.Create;
    L_ValidCyCodeStrList   := TStringList.Create;
    L_GroupInfo   := TStringList.Create;

    AppInfoList := TList.Create;

end;

function TdtmMain.getGroup(nI_CompNdx, nI_Code: Integer): Integer;
var
    nL_RealCode : Integer;
    nL_Result : Integer;
begin
    nL_Result := -1;
    case nI_CompNdx of
     0,1,2,3: //NTY, BEST,SH,CY
       nL_RealCode := nI_Code mod 100;
    end;

    case nI_CompNdx of
     0,1,2,3: //NTY, BEST,SH,CY
      begin
          if (nL_RealCode=1) then
            nL_Result := 1
          else if (nL_RealCode=2) then
            nL_Result := 2
          else if (nL_RealCode>=3) and (nL_RealCode<=14) then
            nL_Result := 3
          else if (nL_RealCode>=15) and (nL_RealCode<=26) then
            nL_Result := 4
          else if (nL_RealCode>=27) and (nL_RealCode<=32) then
            nL_Result := 5
          else if (nL_RealCode>=33) and (nL_RealCode<=40) then
            nL_Result := 6
          else if (nL_RealCode>=41) and (nL_RealCode<=47) then
            nL_Result := 7
          else if (nL_RealCode>=48) and (nL_RealCode<=60) then
            nL_Result := 8
          else if (nL_RealCode>=61) and (nL_RealCode<=79) then
            nL_Result := 9
          else if (nL_RealCode>=80) and (nL_RealCode<=90) then
            nL_Result := 10
          else if (nL_RealCode>=91) and (nL_RealCode<=92) then
            nL_Result := 11
          else if (nL_RealCode>=93) and (nL_RealCode<=99) then
            nL_Result := 12;
      end;

    end;
    result := nL_Result;
end;

procedure TdtmMain.initCDSItems(nI_CompNdx:Integer);
var
    L_StrList : TStringList;
    ii, nL_Group : Integer;
    sL_Content : String;
    sL_Code, sL_Name : String;
    sL_ExePath, sL_ExeFileName : String;

begin
    sL_ExeFileName :=  Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    L_StrList := TStringList.Create;
    case nI_CompNdx of
      0://NTY
        L_StrList.LoadFromFile(sL_ExePath + '\' + 'NTY.txt');
      1://BEST
        L_StrList.LoadFromFile(sL_ExePath + '\' + 'BEST.txt');
      2://SH
        L_StrList.LoadFromFile(sL_ExePath + '\' + 'SH.txt');
      3://CY
        L_StrList.LoadFromFile(sL_ExePath + '\' + 'CY.txt');

    end;
    with cdsActivityCodeStat do
    begin
      close;
      CreateDataSet;
      for ii:=0 to L_StrList.Count-1 do
      begin
        sL_Content := L_StrList.Strings[ii];
        sL_Code := Copy(sL_Content,1,3);
        sL_Name := Copy(sL_Content,5,length(sL_Content)-4);
        nL_Group := getGroup(nI_CompNdx, StrToInt(sL_Code));
        cdsActivityCodeStat.append;
        cdsActivityCodeStat.FieldByName('nGroup').AsInteger := nL_Group;
        cdsActivityCodeStat.FieldByName('sActivityCode').AsString := sL_Code;
        cdsActivityCodeStat.FieldByName('sActivityCodeName').AsString := sL_Name;
        cdsActivityCodeStat.FieldByName('nOccurrence1').AsInteger := 0;
        cdsActivityCodeStat.FieldByName('nOccurrence2').AsInteger := 0;
        cdsActivityCodeStat.FieldByName('nOccurrence3').AsInteger := 0;
        cdsActivityCodeStat.FieldByName('nOccurrence').AsInteger := 0;
        cdsActivityCodeStat.Post;
      end;
    end;
    L_StrList.Free;
end;

procedure TdtmMain.DataModuleDestroy(Sender: TObject);
begin
    ReleaseAppInfoObject;
    AppInfoList.Free;
    L_ValidNtyCodeStrList.free;
    L_ValidBestCodeStrList.free;
    L_ValidShCodeStrList.free;
    L_ValidCyCodeStrList.free;
    L_GroupInfo.Free;
end;

function TdtmMain.checkIfInvalid(nI_CompNdx : Integer; sI_ActivityCode: String): boolean;
var
    nL_Ndx : Integer;
    bL_Result : Boolean;
begin
    //此function用來檢查是否為有效的錯誤資料
    {
    由 sI_ActivityCode 中第一個非零的數字來決定是否納入  invalid code
    }
    bL_Result := false;
    if StrToInt(sI_ActivityCode)=0 then
    begin
      bL_Result := false;
      exit;
    end;

    //down, 若 sI_ActivityCode屬於所應屬的資料區,則必為有效資料
    case nI_CompNdx of
      0: //NTY
        nL_Ndx := L_ValidNTYCodeStrList.IndexOf(sI_ActivityCode);
      1: //BEST
        nL_Ndx := L_ValidBESTCodeStrList.IndexOf(sI_ActivityCode);
      2: //SH
        nL_Ndx := L_ValidSHCodeStrList.IndexOf(sI_ActivityCode);
      3: //CY
        nL_Ndx := L_ValidCYCodeStrList.IndexOf(sI_ActivityCode);

    end;
    if (nL_Ndx<>-1) then
    begin
      bL_Result := true;
      exit;
    end;
    //up, 若 sI_ActivityCode屬於所應屬的資料區,則必為有效資料

    {    //marked   by   Howard,   2002/02/18
    //down, 若 sI_ActivityCode屬於所不應應屬的資料區,則必為無效資料
    case nI_CompNdx of
      0: //NTY
        nL_Ndx := L_ValidLscCodeStrList.IndexOf(sI_ActivityCode);
      1://LSC
        nL_Ndx := L_ValidNTYCodeStrList.IndexOf(sI_ActivityCode);
    end;
    if (nL_Ndx<>-1) then
    begin
      bL_Result := false;
      exit;
    end;
    //up, 若 sI_ActivityCode屬於所不應應屬的資料區,則必為無效資料
    //marked   by   Howard,   2002/02/18
    }
    case nI_CompNdx of
      0: //NTY
       begin
        //down, 屬於NTY的sI_ActivityCode必定是1開頭
        if (Copy(sI_ActivityCode,1,1)='1') then
          bL_Result := true;
        //up, 屬於NTY的sI_ActivityCode必定是1開頭
       end;
      1: //BEST
       begin
        //down, 屬於BEST的sI_ActivityCode必定是2開頭
        if (Copy(sI_ActivityCode,1,1)='2') then
          bL_Result := true;
        //up, 屬於BEST的sI_ActivityCode必定是2開頭
       end;
      2: //SH
       begin
        //down, 屬於SH的sI_ActivityCode必定是3開頭
        if (Copy(sI_ActivityCode,1,1)='3') then
          bL_Result := true;
        //up, 屬於SH的sI_ActivityCode必定是3開頭
       end;
      3: //CY
       begin
        //down, 屬於CY的sI_ActivityCode必定是4開頭
        if (Copy(sI_ActivityCode,1,1)='4') then
          bL_Result := true;
        //up, 屬於Cy的sI_ActivityCode必定是4開頭
       end;
    end;

    result := bL_Result;

end;

procedure TdtmMain.loadTmpCode;
var
    L_StrList : TStringList;
    ii, nL_Group : Integer;
    sL_Content : String;
    sL_Code, sL_Name : String;
    sL_ExePath, sL_ExeFileName : String;

begin
    sL_ExeFileName :=  Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    L_StrList := TStringList.Create;

    //down, load NTY的data...
    L_ValidNtyCodeStrList.Clear;
    L_StrList.LoadFromFile(sL_ExePath + '\' + 'NTY.txt');
    for ii:=0 to L_StrList.Count-1 do
    begin
      sL_Content := L_StrList.Strings[ii];
      sL_Code := Copy(sL_Content,1,3);
//      sL_Name := Copy(sL_Content,5,length(sL_Content)-4);
      L_ValidNtyCodeStrList.Add(sL_Code);
    end;
    //up, load NTY的data...

    //down, load BEST的data...
    L_ValidBestCodeStrList.Clear;
    L_StrList.LoadFromFile(sL_ExePath + '\' + 'BEST.txt');
    for ii:=0 to L_StrList.Count-1 do
    begin
      sL_Content := L_StrList.Strings[ii];
      sL_Code := Copy(sL_Content,1,3);
//      sL_Name := Copy(sL_Content,5,length(sL_Content)-4);
      L_ValidBestCodeStrList.Add(sL_Code);
    end;
    //up, load BEST的data...

    //down, load SH的data...
    L_ValidShCodeStrList.Clear;
    L_StrList.LoadFromFile(sL_ExePath + '\' + 'SH.txt');
    for ii:=0 to L_StrList.Count-1 do
    begin
      sL_Content := L_StrList.Strings[ii];
      sL_Code := Copy(sL_Content,1,3);
//      sL_Name := Copy(sL_Content,5,length(sL_Content)-4);
      L_ValidShCodeStrList.Add(sL_Code);
    end;
    //up, load SH的data...

    //down, load CY的data...
    L_ValidCyCodeStrList.Clear;
    L_StrList.LoadFromFile(sL_ExePath + '\' + 'CY.txt');
    for ii:=0 to L_StrList.Count-1 do
    begin
      sL_Content := L_StrList.Strings[ii];
      sL_Code := Copy(sL_Content,1,3);
//      sL_Name := Copy(sL_Content,5,length(sL_Content)-4);
      L_ValidCyCodeStrList.Add(sL_Code);
    end;
    //up, load CY的data...



    L_StrList.Free;
end;

function TdtmMain.activeAgentPerformanceStat(nI_DataSrc ,nI_CompNdx: Integer; dI_SDate,
  dI_EDate: TDate): boolean;
var
    sL_TableName : String;
    sL_CompCond, sL_CodeName, sL_Code : String;
    nL_Group, nL_Code , nL_Occurrence : Integer;
    bL_HasRecord : boolean;
    wL_Year, wL_Month, wL_Day : Word;
    sL_SDate, sL_EDate, sL_YMD : String;
    nL_AgentLogin : Integer;
begin
    decodeDate(dI_SDate,wL_Year, wL_Month, wL_Day);
//    sL_YMD := Format('%.4d/%.2d/%.2d', [ wL_Year, wL_Month, wL_Day ]);
//    sL_SDate := sL_YMD + ' 00:00:01';

    sL_SDate := Format('%.4d%.2d%.2d', [wL_Year, wL_Month, wL_Day]);

    case nI_DataSrc of
      0: //從Symposium的日報抓data
       begin
        //decodeDate(dI_EDate-1,wL_Year, wL_Month, wL_Day);
        decodeDate( dI_EDate,wL_Year, wL_Month, wL_Day);
        sL_TableName := ' dAgentPerformanceStat ';
       end;
      1: //從Symposium的月報抓data
       begin
        decodeDate(dI_SDate,wL_Year, wL_Month, wL_Day); //因為Symposium的monthly report的date都是該月份的第一天,所以這麼寫程式(用dI_SDate來替代dI_EDate)
        sL_TableName := ' mAgentPerformanceStat ';
       end;
      2: //從Symposium的 interval report抓data
       begin
        decodeDate(dI_EDate,wL_Year, wL_Month, wL_Day);
        sL_TableName := ' iAgentPerformanceStat ';
       end;
    end;
    //sL_YMD := Format('%.4d/%.2d/%.2d', [ wL_Year, wL_Month, wL_Day ]);
    //sL_EDate := sL_YMD + ' 23:59:59';

    sL_EDate := Format('%.4d%.2d%.2d', [wL_Year, wL_Month, wL_Day]);

    sL_CompCond := ' ';

    with qryAgentPerformance do
    begin
      SQL.Clear;
      SQL.Add('select AgentLogin,AgentGivenName,  ');
      SQL.Add('SUM(TalkTime) TalkTime, ');
      SQL.Add('Sum(NotReadyTime) NotReadyTime, ');
      SQL.Add('SUM(CallsOffered) CallsPresented, ');
      SQL.Add('SUM(LoggedInTime) LoggedInTime, ');            
      SQL.Add('SUM(CallsAnswered) CallsAnswered, ');
      SQL.Add('SUM(ShortCallsAnswered) ShortCallsAns ');

      SQL.Add('from ' + sL_TableName);
      //SQL.Add('where Timestamp between :StartDate and :EndDate');
      SQL.Add( Format( 'where convert( varchar(8), Timestamp, 112 ) between ''%s'' and ''%s'' ', [sL_SDate, sL_EDate] ) );
      SQL.Add('group by AgentLogin,AgentGivenName ');
      SQL.Add('order by AgentLogin ');

//      ParamByName('StartDate').AsDateTime := StrToDateTime(sL_SDate);
//      ParamByName('EndDate').AsDateTime := StrToDateTime(sL_EDate);

//      SQL.SaveToFile('c:\xyz.txt');
      Open;
      if (RecordCount>0) then
      begin
        bL_HasRecord := True;
        cdsAgentPerformance.Close;
        cdsAgentPerformance.CreateDataSet;
        while not Eof do
        begin
          nL_AgentLogin := FieldByName('AgentLogin').AsInteger;
          nL_Group := 1;          
          {
          if ((nL_AgentLogin>1000)and(nL_AgentLogin<2000)) then
            nL_Group := 1
          else if((nL_AgentLogin>2000)and(nL_AgentLogin<3000)) then
            nL_Group := 2;
          }
          cdsAgentPerformance.Append;
          cdsAgentPerformance.FieldByName('Group').AsInteger := nL_Group;
          cdsAgentPerformance.FieldByName('AgentLogin').AsString := IntToStr(nL_AgentLogin);
          cdsAgentPerformance.FieldByName('AgentGivenName').AsString := FieldByName('AgentGivenName').AsString;
          cdsAgentPerformance.FieldByName('CallsAnswered').AsInteger := FieldByName('CallsAnswered').AsInteger;
          cdsAgentPerformance.FieldByName('TalkTime').AsInteger := FieldByName('TalkTime').AsInteger;
          cdsAgentPerformance.FieldByName('NotReadyTime').AsInteger := FieldByName('NotReadyTime').AsInteger;
          cdsAgentPerformance.FieldByName('CallsPresented').AsInteger := FieldByName('CallsPresented').AsInteger;
          cdsAgentPerformance.FieldByName('LoggedInTime').AsInteger := FieldByName('LoggedInTime').AsInteger;
          cdsAgentPerformance.FieldByName('ShortCallsAns').AsInteger := FieldByName('ShortCallsAns').AsInteger;
          cdsAgentPerformance.Post;
          Next;
        end;

      end
      else
        bL_HasRecord := false;
      qryAgentPerformance.Close;  
    end;
    result := bL_HasRecord;
end;

function TdtmMain.activeAgentBySkillsetPerformance(nI_CompNdx: Integer;
  dI_SDate, dI_EDate: TDate): boolean;
var
    sL_CompCond, sL_CodeName, sL_Code : String;
    nL_Group, nL_Code , nL_Occurrence : Integer;
    bL_HasRecord : boolean;
    wL_Year, wL_Month, wL_Day : Word;
    sL_SDate, sL_EDate, sL_YMD : String;
begin
    decodeDate(dI_SDate,wL_Year, wL_Month, wL_Day);
    sL_YMD := Format('%.4d/%.2d/%.2d', [ wL_Year, wL_Month, wL_Day ]);
    sL_SDate := sL_YMD + ' 00:00:01';

    decodeDate(dI_EDate,wL_Year, wL_Month, wL_Day);
    sL_YMD := Format('%.4d/%.2d/%.2d', [ wL_Year, wL_Month, wL_Day ]);
    sL_EDate := sL_YMD + ' 23:59:59';

    sL_CompCond := ' ';

    with qryAgentBySkillsetPerformance do
    begin
      SQL.Clear;
      SQL.Add('select AgentSurName, ');
      SQL.Add('SUM(TalkTime) TalkTime, ');
      SQL.Add('Sum(NotReadyTime) NotReadyTime, ');
      SQL.Add('SUM(CallsAnswered) CallsAnswered, ');
      SQL.Add('SUM(TalkTime+NotReadyTime) WorkTime ');      
      SQL.Add('from iAgentBySkillsetStat ');
      SQL.Add('where Timestamp between :StartDate and :EndDate');
      SQL.Add('group by AgentSurName ');
      SQL.Add('order by AgentSurName ');

      ParamByName('StartDate').AsDateTime := StrToDateTime(sL_SDate);
      ParamByName('EndDate').AsDateTime := StrToDateTime(sL_EDate);
//      SQL.SaveToFile('c:\abc.txt');
      Open;
      if (RecordCount>0) then
      begin
        bL_HasRecord := True;
      end
      else
        bL_HasRecord := false;
    end;
    result := bL_HasRecord;
end;

procedure TdtmMain.computeGroupInfo(nI_CompNdx, nI_TotalOccurrence:Integer);
var
    L_Obj : TGroupInfoObj;
    sL_Group, sL_Percent,  sL_ActivityCode   :   String;
    nL_Occurrence,   nL_Ndx,  nL_MemberCount   :   Integer;
    fL_Percent : double;
    nOccurrence1, nOccurrence2, nOccurrence3: Integer;
begin
    L_GroupInfo.Clear;
    with cdsActivityCodeStat do
    begin
      First;
      while not Eof do
      begin
        sL_Group   :=   FieldByName('nGroup').AsString;
        sL_ActivityCode  :=FieldByName('sActivityCode').AsString;
        nL_Occurrence :=   FieldByName('nOccurrence').AsInteger;
        nOccurrence1 :=   FieldByName('nOccurrence1').AsInteger;
        nOccurrence2 :=   FieldByName('nOccurrence2').AsInteger;
        nOccurrence3 :=   FieldByName('nOccurrence3').AsInteger;

        nL_Ndx := L_GroupInfo.IndexOf(sL_Group);
        if nL_Ndx = -1 then
        begin
          nL_MemberCount  :=  1;
          L_Obj := TGroupInfoObj.Create;
          L_Obj.Group := sL_Group;
          L_Obj.TotalOccurrence := nL_Occurrence;
          L_Obj.MemberCount  := nL_MemberCount;
          L_Obj.ActivityCode  :=  sL_ActivityCode;
          
          L_Obj.nOccurrence1 :=  nOccurrence1;
          L_Obj.nOccurrence2 :=  nOccurrence2;
          L_Obj.nOccurrence3 :=  nOccurrence3;

          if (nL_Occurrence<>0) and (nI_TotalOccurrence<>0) then
          begin
            sL_Percent := Format('%8.1f',[(nL_Occurrence/nI_TotalOccurrence)*100]);
            fL_Percent := StrToFloat(sL_Percent);
          end
          else
          begin
            sL_Percent := '0';
            fL_Percent := StrToFloat(sL_Percent);
          end;

          L_Obj.TotalPercent := fL_Percent;
          L_GroupInfo.AddObject(L_Obj.Group,L_Obj);
        end
        else
        begin
          INC(nL_MemberCount);
          (L_GroupInfo.Objects[nL_Ndx]   As TGroupInfoObj).TotalOccurrence :=
            (L_GroupInfo.Objects[nL_Ndx]   As TGroupInfoObj).TotalOccurrence +   nL_Occurrence;
          {}
          TGroupInfoObj(L_GroupInfo.Objects[nL_Ndx]).nOccurrence1 :=
            TGroupInfoObj(L_GroupInfo.Objects[nL_Ndx]).nOccurrence1 + nOccurrence1;

          TGroupInfoObj(L_GroupInfo.Objects[nL_Ndx]).nOccurrence2 :=
            TGroupInfoObj(L_GroupInfo.Objects[nL_Ndx]).nOccurrence2 + nOccurrence2;

          TGroupInfoObj(L_GroupInfo.Objects[nL_Ndx]).nOccurrence3 :=
            TGroupInfoObj(L_GroupInfo.Objects[nL_Ndx]).nOccurrence3 + nOccurrence3;
          {}

          L_Obj.ActivityCode  :=  '';

          if ((L_GroupInfo.Objects[nL_Ndx]   As TGroupInfoObj).TotalOccurrence<>0) and (nI_TotalOccurrence<>0) then
          begin
            sL_Percent := Format('%8.1f',[((L_GroupInfo.Objects[nL_Ndx]   As TGroupInfoObj).TotalOccurrence/nI_TotalOccurrence)*100]);
            fL_Percent := StrToFloat(sL_Percent);
          end
          else
          begin
            sL_Percent := '0';
            fL_Percent := StrToFloat(sL_Percent);
          end;
          (L_GroupInfo.Objects[nL_Ndx] As TGroupInfoObj).MemberCount  := nL_MemberCount;
          (L_GroupInfo.Objects[nL_Ndx] As TGroupInfoObj).TotalPercent := fL_Percent;

        end;
        Next;
      end;
      First;
    end;
end;

procedure TdtmMain.getGroupInfo(sI_Group: String;
  var nI_Occurrence, nI_MemberCount: Integer; var fI_Percent: double;  var  sI_GroupActivityCode  :  String);
var
    nL_Ndx   :   Integer;
begin
    nL_Ndx := L_GroupInfo.IndexOf(sI_Group);
    if   nL_Ndx=-1 then
    begin
      nI_Occurrence   := 0  ;
      fI_Percent   :=   0;
      nI_MemberCount  := 0 ;
      sI_GroupActivityCode  :=  '';
    end
    else
    begin
      nI_Occurrence := (L_GroupInfo.Objects[nL_Ndx]   As TGroupInfoObj).TotalOccurrence;
      fI_Percent :=(L_GroupInfo.Objects[nL_Ndx]   As TGroupInfoObj).TotalPercent;
      nI_MemberCount  := (L_GroupInfo.Objects[nL_Ndx]   As TGroupInfoObj).MemberCount;
      sI_GroupActivityCode  :=(L_GroupInfo.Objects[nL_Ndx]   As TGroupInfoObj).ActivityCode;
    end;
end;

procedure TdtmMain.getGroupInfo2(sI_Group: String; var nIOccurrence1,
  nIOccurrence2, nIOccurrence3: Integer);
var
  aIndex: Integer;
begin
  nIOccurrence1 := 0;
  nIOccurrence2 := 0;
  nIOccurrence3 := 0;
  aIndex := L_GroupInfo.IndexOf( sI_Group );
  if aIndex >=0 then
  begin
    nIOccurrence1 := TGroupInfoObj( L_GroupInfo.Objects[aIndex] ).nOccurrence1;
    nIOccurrence2 := TGroupInfoObj( L_GroupInfo.Objects[aIndex] ).nOccurrence2;
    nIOccurrence3 := TGroupInfoObj( L_GroupInfo.Objects[aIndex] ).nOccurrence3;
  end;
end;

{ ---------------------------------------------------------------------------- }


procedure TdtmMain.Database1Login(Database: TDatabase;
  LoginParams: TStrings);
begin
{
    LoginParams.Clear;
    LoginParams.Add('USERNAME=sysadmin');
    LoginParams.Add('PASSWORD=nortel');
}
end;

procedure TdtmMain.SaveAppObjToFile;
var
  aIndex: Integer;
  aObj: PAppInfo;
  aIniFile: TIniFile;
  aFullName: String;
begin
  aFullName := IncludeTrailingPathDelimiter(
    ExtractFilePath( ParamStr( 0 ) ) ) + APPINFO_FILENAME;
  aIniFile := TIniFile.Create( aFullName );
  try
    for aIndex := 0 to AppInfoList.Count - 1 do
    begin
      aObj := PAppInfo( AppInfoList.Items[aIndex] );
      aIniFile.WriteString( 'ApplicationId', aObj.AppId, aObj.SubAppId );
    end;
    aIniFile.UpdateFile;
  finally
    aIniFile.Free;
  end;
end;

procedure TdtmMain.ReleaseAppInfoObject;
var
  aIndex: Integer;
  aObj: PAppInfo;
begin
  for aIndex := 0 to AppInfoList.Count - 1 do
  begin
    aObj := PAppInfo( AppInfoList.Items[aIndex] );
    Dispose( aObj );
  end;
  AppInfoList.Clear;
end;

end.
