unit frmInvA02_8U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData,
  cxTextEdit, cxCalendar, cxGridLevel, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxClasses, cxControls,
  cxGridCustomView, cxGrid;

type
  TfrmA02_7 = class(TForm)
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
    dsInvLog: TDataSource;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmA02_7: TfrmA02_7;

implementation

{$R *.dfm}

end.
