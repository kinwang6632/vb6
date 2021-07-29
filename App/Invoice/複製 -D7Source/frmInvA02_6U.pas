unit frmInvA02_6U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, cxStyles, cxCustomData, cxGraphics,
  cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData, cxControls,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxClasses, cxGridLevel, cxGrid, Menus,
  cxLookAndFeelPainters, cxButtons, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinscxPCPainter, dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

type
  TfrmInvA02_6 = class(TForm)
    Panel1: TPanel;
    Bevel2: TBevel;
    btnOk: TcxButton;
    btnCancel: TcxButton;
    Panel2: TPanel;
    ResultGrid: TcxGrid;
    LvResult: TcxGridLevel;
    TvResult: TcxGridDBTableView;
    Bevel1: TBevel;
    dsMainInvId: TDataSource;
    TvResultINVID: TcxGridDBColumn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnCancelClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure TvResultDblClick(Sender: TObject);
  private
    { Private declarations }
    FInvId: String;
    FMainInvId: String;
  public
    { Public declarations }
    property InvId: String read FInvId write FInvId;
    property MainInvId: String read FMainInvId write FMainInvId;
  end;

var
  frmInvA02_6: TfrmInvA02_6;

implementation

uses cbUtilis, frmMainU, dtmMainU, dtmMainHU, dtmMainJU, dtmSOU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_6.FormCreate(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_6.FormShow(Sender: TObject);
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    '  SELECT /*+ RULE */                 ' +
    '         A.INVID FROM INV007 A       ' +
    '   WHERE A.IDENTIFYID1 = ''%s''      ' +
    '     AND A.IDENTIFYID2 = %s          ' +
    '     AND A.MAININVID = ''%s''        ' +
    '     AND A.COMPID = ''%s''           ' +
    '     AND A.INVID <> ''%s''           ' +
    '   ORDER BY A.INVID                  ',
    [IDENTIFYID1, IDENTIFYID2, FMainInvId, dtmMain.getCompID, FInvId] );
  dtmMain.adoComm.Open;
  dtmMain.adoComm.First;
  btnOk.Enabled := not ( dtmMain.adoComm.IsEmpty );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_6.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_6.btnCancelClick(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_6.btnOkClick(Sender: TObject);
begin
  FInvId := EmptyStr;
  if not dtmMain.adoComm.IsEmpty then
  begin
    FInvId := dtmMain.adoComm.FieldByName( 'INVID' ).AsString;
    Self.ModalResult := mrOk;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_6.TvResultDblClick(Sender: TObject);
begin
  btnOkClick( btnOk );
end;

{ ---------------------------------------------------------------------------- }

end.
