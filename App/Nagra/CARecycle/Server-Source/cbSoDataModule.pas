unit cbSoDataModule;

interface

uses
  SysUtils, Classes, DB, ADODB, ActiveX;


  { SoConnection 指的是連線到那一區抓 CARECYCLE 的 Table 資料 }
  { CAConnection 指的是把須要傳 CA 指令寫到那一個 COM 區 }

type
  TSoDataModule = class(TDataModule)
    SoConnection: TADOConnection;
    SoDataReader: TADOQuery;
    SoDataWriter: TADOQuery;
    CAConnection: TADOConnection;
    CADataReader: TADOQuery;
    CADataWriter: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SoDataModule: TSoDataModule;

implementation


{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.DataModuleCreate(Sender: TObject);
begin
  CoInitialize( nil );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDataModule.DataModuleDestroy(Sender: TObject);
begin
  CoUninitialize;
end;

{ ---------------------------------------------------------------------------- }

end.
