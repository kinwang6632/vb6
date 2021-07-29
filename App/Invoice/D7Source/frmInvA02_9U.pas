unit frmInvA02_9U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, DBClient, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxCalendar, Menus, cxLookAndFeelPainters, cxButtons;

type
  TfrmInvA02_9 = class(TForm)
    Panel1: TPanel;
    Label3: TLabel;
    txtStartDate: TcxDateEdit;
    txtEndDate: TcxDateEdit;
    btnSave: TcxButton;
    btnCancel: TcxButton;
    Label1: TLabel;
    procedure FormShow(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
    FDataSet: TClientDataSet;
    FBillId:String;
    FBillIDItemNo:String;
    FSeq:String;
  public
    { Public declarations }
    constructor Create(aDataSet: TClientDataSet); reintroduce;
  end;

var
  frmInvA02_9: TfrmInvA02_9;

implementation

{$R *.dfm}

{ TfrmInvA02_9 }

constructor TfrmInvA02_9.Create(aDataSet: TClientDataSet);
begin
  inherited Create( Application );
  FDataSet := aDataSet;
//  FReader := DBController.CodeReader;
end;

procedure TfrmInvA02_9.FormShow(Sender: TObject);
begin
  txtStartDate.Text := FDataSet.FieldByName('StartDate').Text;
  txtEndDate.Text := FDataSet.FieldByName('EndDate').Text;
end;

procedure TfrmInvA02_9.btnSaveClick(Sender: TObject);
begin
  FDataSet.Edit;
  FDataSet.FieldByName('StartDate').AsDateTime := txtStartDate.Date;
  FDataSet.FieldByName('EndDate').AsDateTime := txtEndDate.Date;
  FDataSet.Post;
  Self.Close;
end;

procedure TfrmInvA02_9.btnCancelClick(Sender: TObject);
begin
  Self.Close;
end;

end.
