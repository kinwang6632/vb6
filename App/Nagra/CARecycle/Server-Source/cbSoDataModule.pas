unit cbSoDataModule;

interface

uses
  SysUtils, Classes, DB, ADODB, ActiveX;


  { SoConnection �����O�s�u�쨺�@�ϧ� CARECYCLE �� Table ��� }
  { CAConnection �����O�⶷�n�� CA ���O�g�쨺�@�� COM �� }

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
