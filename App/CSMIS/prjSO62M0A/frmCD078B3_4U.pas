unit frmCD078B3_4U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls,
  cbDBController, cbDataLookup, frmCD078B1U;

type
  TfrmCD078B3_4 = class(TForm)
    CItem: TDataLookup;
    Label1: TLabel;
    Bevel1: TBevel;
    btnSave: TButton;
    Button1: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure CItemCodeNoPropertiesChange(Sender: TObject);
    procedure CItemCodeNamePropertiesChange(Sender: TObject);
    procedure CItemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
  private
    { Private declarations }
    FCItemCode: String;
    FServiceType: String;
    FBPCode: String;
    FAlreadySetCItemValue: String;
  public
    { Public declarations }
    property BPCode: String read FBPCode write FBPCode;
    property ServiceType: String read FServiceType write FServiceType;
    property CItemCode: String read FCItemCode write FCItemCode;
    property AlreadySetCItemValue: String read FAlreadySetCItemValue write FAlreadySetCItemValue;
  end;

var
  frmCD078B3_4: TfrmCD078B3_4;

implementation

{$R *.dfm}

uses cbUtilis;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_4.FormCreate(Sender: TObject);
begin
  CItem.Initializa;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_4.FormDestroy(Sender: TObject);
begin
  CItem.Finaliza;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_4.CItemCodeNoPropertiesChange(Sender: TObject);
begin
  CItem.CodeNoPropertiesChange( Sender );
  btnSave.Enabled := ( CItem.Value <> EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_4.CItemCodeNamePropertiesChange(Sender: TObject);
begin
  CItem.CodeNamePropertiesChange( Sender );
  btnSave.Enabled := ( CItem.Value <> EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_4.CItemCodeNamePropertiesInitPopup(Sender: TObject);
var
  aCItems: String;
begin
  Screen.Cursor := crSQLWait;
  try
    CItem.SQL.Text := Format(
      ' SELECT A.CODENO, A.DESCRIPTION      ' +
      '  FROM %s.CD019 A                    ' +
      ' WHERE A.PERIODFLAG = ''1''          ' +
      '   AND A.STOPFLAG = ''0''            ' +
      '   AND A.SERVICETYPE = ''%s''        ',
      [DBController.LoginInfo.DbAccount, FServiceType] );
    if ( FAlreadySetCItemValue <> EmptyStr ) then
    begin
      aCItems := QuotedValue( FAlreadySetCItemValue );
      CItem.SQL.Text := CItem.SQL.Text +
        ' AND A.CODENO NOT IN ( ' + aCItems + ' ) ';
    end;
    {}
    { #4407 ?p?G?O???~?H???h?W?[?P?_ MASTERPRODUCT <> 1  By Kin 2009/03/06 }
    { ???M???????D By Kin 2009/03/06 }
    if ( frmCD078B1.KindFunction = '1' ) then
    begin
      CItem.SQL.Text := CItem.SQL.Text + ' AND Nvl(A.REFNO,0) = 2 ';
    end;
    {}
    CItem.SQL.Text := CItem.SQL.Text +
      '  ORDER BY A.CODENO             ';
    CItem.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3_4.btnSaveClick(Sender: TObject);
begin
  if ( CItem.Value <> EmptyStr ) and ( CItem.ValueName <> EmptyStr ) then
    Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

end.
