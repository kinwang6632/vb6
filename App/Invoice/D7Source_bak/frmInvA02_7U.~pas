unit frmInvA02_7U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData,
  cxTextEdit, cxCalendar, ADODB, cxGridLevel, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxClasses, cxControls,
  cxGridCustomView, cxGrid;

type
  TfrmInvA02_7 = class(TForm)
    cxGrid: TcxGrid;
    cxGridBrow: TcxGridDBTableView;
    cxGridBrowType: TcxGridDBColumn;
    cxGridBrowCURRENTSTATE: TcxGridDBColumn;
    cxGridBrowType2: TcxGridDBColumn;
    cxGridBrowSendType: TcxGridDBColumn;
    cxGridBrowSendStats: TcxGridDBColumn;
    cxGridBrowSendCnt: TcxGridDBColumn;
    cxGridBrowReturnCode: TcxGridDBColumn;
    cxGridBrowReturnMsg: TcxGridDBColumn;
    cxGridLevel: TcxGridLevel;
    DataSource1: TDataSource;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure cxGridBrowTypeGetDisplayText(Sender: TcxCustomGridTableItem;
      ARecord: TcxCustomGridRecord; var AText: String);
    procedure cxGridBrowSendStatsGetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: String);
  private
    { Private declarations }
    FInvId : String;
    FCompId : String;
    function GetLogData(aInvId,aCompId : String) : Boolean;

  public
    constructor Create(aInvId,aCompId: String); reintroduce;
    { Public declarations }
  end;

var
  frmInvA02_7: TfrmInvA02_7;

implementation
uses cbUtilis, Uotheru, frmMainU, dtmMainU, dtmMainHU, dtmMainJU, dtmSOU,
     StdConvs;
{$R *.dfm}

procedure TfrmInvA02_7.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  try
    ADOConnection1.Close;
  finally

  end;
end;

function TfrmInvA02_7.GetLogData(aInvId, aCompId: String): Boolean;
  var aSQL : String;

begin
  try


    ADOConnection1.Connected := False;
    ADOConnection1.ConnectionString :=Format(
     'Provider=MSDAORA.1;Password=%s;User ID=%s;' +
     'Data Source=%s;Persist Security Info=True',
     [dtmMain.GetEInvPassword, dtmMain.GetEInvUser, dtmMain.GetEInvDb] );
    ADOConnection1.Connected := True;

  except
    on E: Exception do
    begin
      MessageDlg( Format( '【傳送通知】資料區連結失敗, 訊息:%s', [E.Message] ),mtWarning,[mbOK],0);
      Exit;
    end;
  end;


  {INV036}
  aSQL := aSQL + Format('SELECT 1 TYPE,MAIL TYPE2,BOOKINGTIME SENDTIME,        ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,CURRENTSTATENAME,                   ' +
      ' RETURNCODE,RETURNMSG, INV_SENDSTATUS.CURRENTSTATE                      ' +
      ' FROM INV036ERR , INV_SENDSTATUS                                        ' +
      ' WHERE INVID(+) = ''%s''                                                ' +
      ' AND INV036ERR.CURRENTSTATE = INV_SENDSTATUS.CURRENTSTATE(+)            ' +
      ' AND INV_SENDSTATUS.SENDTYPE(+) = 1                                     ' +
      ' AND COMPID(+) = ''%s''                                                 ' ,
      [ aInvId,aCompId ] );
  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 1 TYPE,MAIL TYPE2, BOOKINGTIME SENDTIME,       ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,CURRENTSTATENAME,                   ' +
      ' RETURNCODE,RETURNMSG,INV_SENDSTATUS.CURRENTSTATE                       ' +
      ' FROM INV036LOG , INV_SENDSTATUS                                        ' +
      ' WHERE INVID(+) = ''%s''                                                ' +
      ' AND INV036LOG.CURRENTSTATE = INV_SENDSTATUS.CURRENTSTATE(+)            ' +
      ' AND INV_SENDSTATUS.SENDTYPE(+) = 1                                     ' +
      ' AND COMPID(+) = ''%s''                                                 ' ,
      [ aInvId,aCompId ] );

  {INV037}

  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 2 TYPE,CONTMOBILE TYPE2, BOOKINGTIME SENDTIME,  ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,CURRENTSTATENAME,                    ' +
      ' RETURNCODE,RETURNMSG,INV_SENDSTATUS.CURRENTSTATE                        ' +
      ' FROM INV037ERR , INV_SENDSTATUS                                         ' +
      ' WHERE INVID(+) = ''%s''                                                 ' +
      ' AND INV037ERR.CURRENTSTATE = INV_SENDSTATUS.CURRENTSTATE(+)             ' +
      ' AND INV_SENDSTATUS.SENDTYPE(+) = 2                                      ' +
      ' AND COMPID(+) = ''%s''                                                  ' ,
      [ aInvId,aCompId ] );

  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 2 TYPE,CONTMOBILE TYPE2, BOOKINGTIME SENDTIME,    ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,CURRENTSTATENAME,                      ' +
      ' RETURNCODE,RETURNMSG, INV_SENDSTATUS.CURRENTSTATE                         ' +
      ' FROM INV037LOG , INV_SENDSTATUS                                           ' +
      ' WHERE INVID(+) = ''%s''                                                   ' +
      ' AND INV037LOG.CURRENTSTATE = INV_SENDSTATUS.CURRENTSTATE(+)               ' +
      ' AND INV_SENDSTATUS.SENDTYPE(+) = 2                                        ' +
      ' AND COMPID(+) = ''%s''                                                    ' ,
      [ aInvId,aCompId ] );

  {INV038}
  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 3 TYPE,FACISNO TYPE2, BOOKINGTIME SENDTIME,       ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT, NULL CURRENTSTATENAME,                ' +
      ' RETURNCODE,RETURNMSG,NULL CURRENTSTATE                                    ' +
      ' FROM INV038ERR WHERE INVID = ''%s''                                       ' +
      ' AND COMPID = ''%s''                                                       ' ,
      [ aInvId,aCompId ] );

  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 3 TYPE,FACISNO TYPE2, BOOKINGTIME SENDTIME,       ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT, NULL CURRENTSTATENAME,                ' +
      ' RETURNCODE,RETURNMSG,NULL CURRENTSTATE                                    ' +
      ' FROM INV038LOG WHERE INVID = ''%s''                                       ' +
      ' AND COMPID = ''%s''                                                       ' ,
      [ aInvId,aCompId ] );

  {INV039}
  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 4 TYPE,FACISNO TYPE2,BOOKINGTIME SENDTIME,        ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,NULL CURRENTSTATENAME,                 ' +
      ' RETURNCODE,RETURNMSG,NULL CURRENTSTATE                                    ' +
      ' FROM INV039ERR WHERE INVID = ''%s''                                       ' +
      ' AND COMPID = ''%s''                                                       ' ,
      [ aInvId,aCompId ] );

  aSQL := aSQL + ' UNION ALL ';
  aSQL := aSQL + Format('SELECT 4 TYPE,FACISNO TYPE2,BOOKINGTIME SENDTIME,        ' +
      ' NVL(SENDSTATE,0) SENDSTATE,SENDCNT,NULL CURRENTSTATENAME,                 ' +
      ' RETURNCODE,RETURNMSG,NULL CURRENTSTATE                                    ' +
      ' FROM INV039LOG WHERE INVID = ''%s''                                       ' +
      ' AND COMPID = ''%s''                                                       ' ,
      [ aInvId,aCompId ] );
  aSQL := 'SELECT A.* FROM ( ' + aSQL + ') A ORDER BY TYPE ASC,SENDTIME DESC';
  ADOQuery1.Close;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Text := aSQL;
  ADOQuery1.Open;
  Result := True;
end;

constructor TfrmInvA02_7.Create(aInvId, aCompId: String);
begin
  inherited Create( Application );
  FInvId := aInvId;
  FCompId := aCompId;
  GetLogData( FInvId,FCompId );

end;

procedure TfrmInvA02_7.FormShow(Sender: TObject);
begin
//  GetLogData( FInvId,FCompId );
end;

procedure TfrmInvA02_7.cxGridBrowTypeGetDisplayText(
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

procedure TfrmInvA02_7.cxGridBrowSendStatsGetDisplayText(
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
