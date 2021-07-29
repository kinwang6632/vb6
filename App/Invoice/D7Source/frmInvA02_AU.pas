unit frmInvA02_AU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, cxLabel, cxControls, cxContainer,
  cxEdit, cxTextEdit, cxDBEdit, cxDBLabel, DB, DBClient, Provider,
  cxStyles, dxSkinscxPCPainter, cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, cxDBData, cxGridLevel, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxClasses, cxGridCustomView, cxGrid;

type
  TfrmInvA02_A = class(TForm)
    pnl1: TPanel;
    dspINV046: TDataSetProvider;
    cdsMaster: TClientDataSet;
    dsMaster: TDataSource;
    cxGrid1: TcxGrid;
    cxGridDBTableView1: TcxGridDBTableView;
    cxPrintTime: TcxGridDBColumn;
    csPrintEn: TcxGridDBColumn;
    csReasonCode: TcxGridDBColumn;
    cxReasonName: TcxGridDBColumn;
    cxGridLevel1: TcxGridLevel;
  private
    { Private declarations }
    FInvId: String;
    FCompId:String;
    function GetDataLog(aInvId,aCompId : String) : Boolean;
  public
    { Public declarations }
    constructor Create(aInvId,aCompId: String); reintroduce;
  end;

var
  frmInvA02_A: TfrmInvA02_A;

implementation
uses cbUtilis, Uotheru, frmMainU, dtmMainU, dtmMainHU, dtmMainJU, dtmSOU;

{$R *.dfm}

{ TfrmInvA02_A }

constructor TfrmInvA02_A.Create(aInvId, aCompId: String);
begin
  inherited Create( Application );
  FInvId := aInvId;
  FCompId := aCompId;
  GetDataLog(aInvId,aCompId);
end;

function TfrmInvA02_A.GetDataLog(aInvId, aCompId: String): Boolean;
  var aSQL :String;
begin
  aSQL :=Format( 'SELECT * FROM INV046              ' +
        ' WHERE INVID = ''%s''                      ',
        [aInvId]);
  try
    Application.ProcessMessages;
//    cdsMaster.AfterScroll := nil;
    try
      cdsMaster.Close;
      cdsMaster.CommandText := aSQL;
      cdsMaster.Open;
//      SetMasterFieldState;
      Application.ProcessMessages;
    finally
//      cdsMaster.AfterScroll := aEvent;
    end;
    Application.ProcessMessages;
  finally
    Screen.Cursor := crDefault;
  end;
  Result := True;
end;

end.
