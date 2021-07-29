unit frmCompDataU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frmSTabMaintainU, DB, Grids, DBGrids, ExtCtrls, StdCtrls,
  Buttons, DBCtrls, Mask;

type
  TfrmCompData = class(TfrmSTabMaintain)
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    DBRadioGroup1: TDBRadioGroup;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCompData: TfrmCompData;

implementation

uses dtmCompDataU, dtmSTabMaintainu, frmMainU;

{$R *.dfm}

procedure TfrmCompData.FormCreate(Sender: TObject);
begin
  inherited;
    DefaultEditFieldName := 'CODENO';
    SetQueryCond ('公司代碼','CODENO',1);
    SetQueryCond ('公司名稱','DESCRIPTION',2);

    NoneDupFieldName.Add('CODENO');
    NoneDupFieldNameType.Add(INT_TYPE);    
    UnEditedFieldNames.Add('CODENO');
    SetDataModule(dtmCompData);
    ActionTabName := 'EMC_EBT010';

    self.Caption := frmMain.getCaption('公司別代碼資料維護');        
end;

end.
