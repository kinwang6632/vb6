unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, DBClient;

type
  TdtmMain = class(TDataModule)
    cdsParam: TClientDataSet;
    cdsParamsServerIP2: TStringField;
    cdsParamnSPortNo2: TIntegerField;
    cdsParamnRPortNo2: TIntegerField;
    cdsParamsSysName2: TStringField;
    cdsParamsSysVersion2: TStringField;
    cdsParamsLogPath: TStringField;
    cdsParamnTimeOut2: TIntegerField;
    cdsParambCommandLog2: TBooleanField;
    cdsParambResponseLog: TBooleanField;
    cdsParamdUptTime2: TDateTimeField;
    cdsParamsUptName2: TStringField;
    cdsParambShowUI2: TBooleanField;
    cdsParamnVersion: TIntegerField;
    cdsParamsSecuritytype: TStringField;
    cdsParamnToID: TIntegerField;
    cdsParamnConnectionID: TIntegerField;
    cdsParamnForwardTolerance: TIntegerField;
    cdsParamnBackwardTolerance: TIntegerField;
    cdsParamnCurrency: TIntegerField;
    cdsParamnFromID: TIntegerField;
    cdsParamnTimeZoneOffset: TIntegerField;
    cdsParamsCountryNumber: TStringField;
    cdsParamnTimeZoneOffset2: TIntegerField;
    cdsUser: TClientDataSet;
    cdsUsersID: TStringField;
    cdsUsersName: TStringField;
    cdsUsersPassword: TStringField;
    cdsParamsKey: TStringField;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  dtmMain: TdtmMain;

implementation

{$R *.dfm}

end.
