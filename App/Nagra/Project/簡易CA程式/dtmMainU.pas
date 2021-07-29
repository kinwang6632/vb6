unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, Forms, Dialogs, DBClient, Variants;

type
  TdtmMain = class(TDataModule)
    cdsUser: TClientDataSet;
    cdsUsersID: TStringField;
    cdsUsersName: TStringField;
    cdsUsersPassword: TStringField;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }

    procedure ActiveCDS(I_CDS:TClientDataSet; sI_RefFileName:String; sI_DefaultStatus:String);
  public
    { Public declarations }
    function isCorrectID(sI_ID, sI_Passwd:String):boolean;
  end;

var
  dtmMain: TdtmMain;

implementation

uses ConstObjU;


{$R *.dfm}









procedure TdtmMain.ActiveCDS(I_CDS: TClientDataSet; sI_RefFileName,
  sI_DefaultStatus: String);
var
    sL_Path : String;
begin
//    SetCurrentDir(ExtractFilePath(Application.ExeName));

    sL_Path := ExtractFilePath(Application.ExeName);

    if (I_CDS.State = dsInactive) then
      I_CDS.CreateDataSet;
    if (FileExists(sL_Path + sI_RefFileName)) then
      I_CDS.LoadFromFile(sL_Path + sI_RefFileName);

    if (I_CDS.RecordCount=0) then
      I_CDS.append
    else
      I_CDS.edit;

    if (sI_DefaultStatus='E') then //edit mode
    begin
    end
    else if (sI_DefaultStatus='B') then//browse mode
    begin
      I_CDS.Post;
    end;

end;

procedure TdtmMain.DataModuleCreate(Sender: TObject);
begin
    ActiveCDS(dtmMain.cdsUser, USER_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
end;

function TdtmMain.isCorrectID(sI_ID, sI_Passwd: String): boolean;
var
    bL_Result : boolean;
begin
    with cdsUser do
    begin
      bL_Result := Locate('sID;sPassword', VarArrayOf([sI_ID, sI_Passwd]), [loPartialKey]);
    end;

    result := bL_Result;
end;

end.
