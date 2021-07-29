unit frmInvA02_BU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData,
  cxTextEdit, cxCalendar, ADODB, cxGridLevel, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxClasses, cxControls,
  cxGridCustomView, cxGrid, Provider, DBClient;

type
  TfrmInvA02_B = class(TForm)
    dsMaster: TDataSource;
    cxGrid1: TcxGrid;
    cxGridDBTableView1: TcxGridDBTableView;
    cxUploadType: TcxGridDBColumn;
    cxUploadTime: TcxGridDBColumn;
    cxUploadSource: TcxGridDBColumn;
    cxGridLevel1: TcxGridLevel;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    FDataSet: TClientDataSet;
    function GetLogData : Boolean;
  public
    { Public declarations }
    constructor Create(aDataSet:TClientDataSet); reintroduce;
  end;

var
  frmInvA02_B: TfrmInvA02_B;

implementation
uses cbUtilis, Uotheru, frmMainU, dtmMainU, dtmMainHU, dtmMainJU, dtmSOU,
     StdConvs;
{$R *.dfm}

{ TfrmInvA02_B }

constructor TfrmInvA02_B.Create(aDataSet:TClientDataSet);
begin
  inherited Create( Application );
  FDataSet := aDataSet;
  dsMaster.DataSet := FDataSet;
//  cdsMaster := aDataSet;
  {
  if not GetLogData then
  begin
    WarningMsg( '查無任何上傳記錄' );
    Close;
  end;
  }
end;

function TfrmInvA02_B.GetLogData: Boolean;
 var
   aSQL:string;
begin
  {
  aSQL := Format('select case uploadtype         ' +
        ' when 1 then ''C0401''                  ' +
        ' when 2 then ''C0501''                  ' +
        ' when 3 then ''D0401''                  ' +
        ' when 4 then ''D0501''                  ' +
        ' when 5 then ''C0701''                  ' +
        ' when 6 then ''E0402''                  ' +
        ' else ''未知'' end uploadtype,          ' +
        ' uploadtime,                            ' +
        ' Decode(uploadSource,1,''自動'',''手動'') uploadsource ' +
        ' from inv048                            ' +
        ' where invid=''%s''                     ' +
        ' and compid = ''%s''                    ' +
        ' order by uploadtime                    ',
        [FInvId,FCompId]);
  cdsMaster.Close;
  cdsMaster.CommandText := aSQL;
  cdsMaster.Open;
  if cdsMaster.RecordCount = 0 then
  begin

    Result := False;
  end else
  begin
    Result := True;
  end;
  }  
end;

procedure TfrmInvA02_B.FormShow(Sender: TObject);
begin
// ShowMessage(  IntToStr( cdsMaster.RecordCount ) );
 
end;

end.
