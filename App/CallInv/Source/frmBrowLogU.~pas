unit frmBrowLogU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData,
  cxTextEdit, cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxClasses, cxControls, cxGridCustomView, cxGrid,
  ExtCtrls, cxCalendar;

type
  TfrmLog = class(TForm)
    Panel1: TPanel;
    cxGrid: TcxGrid;
    cxGridBrow: TcxGridDBTableView;
    cxGridLevel: TcxGridLevel;
    dsInvLog: TDataSource;
    cxGridBrowType: TcxGridDBColumn;
    cxGridBrowSendType: TcxGridDBColumn;
    cxGridBrowSendStats: TcxGridDBColumn;
    cxGridBrowSendCnt: TcxGridDBColumn;
    cxGridBrowReturnCode: TcxGridDBColumn;
    cxGridBrowReturnMsg: TcxGridDBColumn;
    cxGridBrowType2: TcxGridDBColumn;
    cxGridBrowCURRENTSTATE: TcxGridDBColumn;
    cxGridBrowREADTIME: TcxGridDBColumn;
    procedure cxGridBrowTypeGetDisplayText(Sender: TcxCustomGridTableItem;
      ARecord: TcxCustomGridRecord; var AText: String);
    procedure cxGridBrowSendStatsGetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: String);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLog: TfrmLog;

implementation
  uses CommU,dtmCommU;
{$R *.dfm}

procedure TfrmLog.cxGridBrowTypeGetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: String);
begin
  if  AText = '1' then
    AText := 'E-MAIL';
  if AText = '2' then
    AText :='簡訊';
  if AText = '3' then
    AText :='TVMAIL';
  if AText = '4' then
    AText := 'CM導流';
end;

procedure TfrmLog.cxGridBrowSendStatsGetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: String);
begin
  //    0:未發  1:已發OK  2:已發失敗  3:重送
  if AText = '0' then
    AText := '未發';
  if AText = '1' then
    AText := '已發OK';
  if AText = '2' then
    AText := '已發失敗';
  if AText = '3' then
    AText := '重送';



end;

end.
