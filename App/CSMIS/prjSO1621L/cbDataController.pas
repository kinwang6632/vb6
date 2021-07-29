unit cbDataController;

interface

uses
  SysUtils, Classes, DB, ADODB;

type

  PLoginInfo = ^TLoginInfo;
  TLoginInfo = record
    IsSuperVisor: Boolean;
    UserId: String;
    UserName: String;
    CompCode: String;
    CompName: String;
    DbAccount: String;
    DbPassword: String;
    DbAliase: String;
    ServiceType: String;
  end;


  TDBController = class(TDataModule)
    DataConnection: TADOConnection;
    DataReader: TADOQuery;
    DataWriter: TADOQuery;
    CodeReader: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FLoginInfo: PLoginInfo;
    function GetConnected: Boolean;
    procedure SetConnected(const Value: Boolean);
  public
    { Public declarations }
    property LoginInfo: PLoginInfo read FLoginInfo;
    property Connected: Boolean read GetConnected write SetConnected;

  end;

var
  DBController: TDBController;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TDBController.DataModuleCreate(Sender: TObject);
begin
  New( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }

procedure TDBController.DataModuleDestroy(Sender: TObject);
begin
  Dispose( FLoginInfo );
end;

{ ---------------------------------------------------------------------------- }



function TDBController.GetConnected: Boolean;
begin
  Result := DataConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TDBController.SetConnected(const Value: Boolean);
begin
  if ( Value ) then
  begin
    DataConnection.ConnectionString := Format(
     'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
      [FLoginInfo.DbPassword, FLoginInfo.DbAccount, FLoginInfo.DbAliase] );
    try
      DataConnection.Connected := True;
    except
      on E: Exception do ErrorMsg( E.Message );
    end;
  end else
  begin
    DataConnection.Connected := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
