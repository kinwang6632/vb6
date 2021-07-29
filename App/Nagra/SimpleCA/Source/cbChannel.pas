unit cbChannel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, cbMain,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxStyles, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, DB, cxDBData, cxGridLevel,
  cxClasses, cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGrid, ADODB;

type
  TfmChannel = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Bevel1: TBevel;
    ChGridView: TcxGridDBTableView;
    ChGridLevel: TcxGridLevel;
    ChGrid: TcxGrid;
    Bevel2: TBevel;
    CD024Reader: TADOQuery;
    btnOK: TButton;
    btnCancel: TButton;
    CD024DataSource: TDataSource;
    ViewCODENO: TcxGridDBColumn;
    ViewDESCRIPTION: TcxGridDBColumn;
    ViewCHANNELID: TcxGridDBColumn;
    Bevel3: TBevel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmChannel: TfmChannel;

implementation

{$R *.dfm}

uses cbUtilis;

procedure TfmChannel.FormCreate(Sender: TObject);
begin
  btnOK.Enabled := False;
  CD024Reader.Close;
  CD024Reader.SQL.Text := ' SELECT * FROM CD024 ORDER BY TO_NUMBER( CHANNELID  ) ';
  try
    CD024Reader.Open;
    btnOK.Enabled := not CD024Reader.IsEmpty;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '列出頻道代碼有誤, 原因:%s。', [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmChannel.FormShow(Sender: TObject);
begin
  ViewCODENO.Visible := False;
  ViewCHANNELID.Visible := False;
  if not CD024Reader.IsEmpty then ChGridView.Controller.SelectAll;
  ChGrid.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmChannel.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfmChannel.FormDestroy(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

end.
