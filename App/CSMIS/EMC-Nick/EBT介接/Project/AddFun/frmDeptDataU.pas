unit frmDeptDataU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frmSTabMaintainU, DB, Grids, DBGrids, ExtCtrls, StdCtrls,
  Buttons, DBCtrls, Mask;

type
  TfrmDeptData = class(TfrmSTabMaintain)
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
  frmDeptData: TfrmDeptData;

implementation

uses dtmDeptDataU, dtmSTabMaintainu, frmMainU;

{$R *.dfm}

procedure TfrmDeptData.FormCreate(Sender: TObject);
begin
  inherited;
    DefaultEditFieldName := 'CODENO';
    SetQueryCond ('部門代碼','CODENO',1);
    SetQueryCond ('部門名稱','DESCRIPTION',2);

    NoneDupFieldName.Add('CODENO');
    NoneDupFieldNameType.Add(INT_TYPE);    
    UnEditedFieldNames.Add('CODENO');
    SetDataModule(dtmDeptData);
    ActionTabName := 'EMC_EBT011';

    self.Caption := frmMain.getCaption('部門別代碼資料維護');    
end;

end.
