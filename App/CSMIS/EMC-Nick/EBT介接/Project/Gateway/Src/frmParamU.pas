unit frmParamU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, DBCtrls, Mask, DB;

type
  TfrmParam = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    btnExit: TBitBtn;
    btnEdit: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    dsrParam: TDataSource;
    pnlEdit: TPanel;
    RadioGroup1: TRadioGroup;
    DBEdit1: TDBEdit;
    StaticText3: TStaticText;
    RadioGroup2: TRadioGroup;
    DBEdit2: TDBEdit;
    StaticText2: TStaticText;
    StaticText1: TStaticText;
    DBCheckBox1: TDBCheckBox;
    DBCheckBox2: TDBCheckBox;
    DBCheckBox3: TDBCheckBox;
    procedure BitBtn4Click(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
  private
    { Private declarations }
    procedure ChangeBtnStatus;    
  public
    { Public declarations }
  end;

var
  frmParam: TfrmParam;

implementation

uses frmMainU, dtmMainU, ConstObjU;

{$R *.dfm}

procedure TfrmParam.BitBtn4Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmParam.btnExitClick(Sender: TObject);
begin
    MessageDlg('э把计斗穝币笆祘Α,穝把计穦ネ!',mtInformation,[mbOK],0);
    Close;
end;

procedure TfrmParam.FormShow(Sender: TObject);
begin
    Caption := frmMain.getFrmCaption('╰参把计砞﹚');
    ChangeBtnStatus;
end;

procedure TfrmParam.btnSaveClick(Sender: TObject);
begin
    dtmMain.cdsParam.Post;
    ChangeBtnStatus;
    dtmMain.SaveCDS(dtmMain.cdsParam, PARAM_INFO_FILENAME);
    dtmMain.setTirggerStatus;
end;

procedure TfrmParam.ChangeBtnStatus;
begin
     with dtmMain.cdsParam do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          pnlEdit.Enabled := TRUE;
          btnCancel.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
       end
       else
       begin
          pnlEdit.Enabled := FALSE;
          if dtmMain.cdsParam.RecordCount>0 then
          begin
            btnEdit.Enabled := TRUE;
          end
          else
          begin
            btnEdit.Enabled := FALSE;
          end;
          btnCancel.Enabled := FALSE;
          btnSave.Enabled := FALSE;
          btnExit.Enabled := TRUE;
       end;
     end;

end;

procedure TfrmParam.btnCancelClick(Sender: TObject);
begin
      dtmMain.cdsParam.Cancel;
      ChangeBtnStatus;

end;

procedure TfrmParam.btnEditClick(Sender: TObject);
begin
    dtmMain.cdsParam.Edit;
    ChangeBtnStatus;

end;

end.
