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
    SetQueryCond ('���q�N�X','CODENO',1);
    SetQueryCond ('���q�W��','DESCRIPTION',2);

    NoneDupFieldName.Add('CODENO');
    NoneDupFieldNameType.Add(INT_TYPE);    
    UnEditedFieldNames.Add('CODENO');
    SetDataModule(dtmCompData);
    ActionTabName := 'EMC_EBT010';

    self.Caption := frmMain.getCaption('���q�O�N�X��ƺ��@');        
end;

end.
