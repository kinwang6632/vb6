unit dtmSTabMaintainu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBTables, Db, DBClient, Provider, ADODB;
const
    STR_TYPE = '1';
    INT_TYPE = '2';    
type
  TdtmSTabMaintain = class(TDataModule)
    ADOQuery1: TADOQuery;
    DataSetProvider1: TDataSetProvider;
    ClientDataSet1: TClientDataSet;
    qryCheckDup1: TADOQuery;
    procedure dtmSTabMaintainCreate(Sender: TObject);
    procedure dtmSTabMaintainDestroy(Sender: TObject);
    procedure SingleTabDBLogin(Database: TDatabase; LoginParams: TStrings);
  private
    { Private declarations }
    FMstQry : TDataSet ;
    FMasterQuery: TQuery;
    function GetMasterDataSet: TDataSet;
    procedure SetMasterQuery(const Value: TQuery);
  public
    { Public declarations }
    ConstFieldNameStrList : TStringList;
    ConstFieldValueStrList : TStringList;
    function CheckDup(sI_TableName, sI_FieldName, sI_Value, sI_FieldType : String):Boolean;
    function FieldValue(sI_FieldName:String):Variant;    
    property MasterDataSet : TDataSet read  GetMasterDataSet ;
    property MasterQuery : TQuery read FMasterQuery write SetMasterQuery;
  protected

    procedure SetMasterData(var MstQry: TDataSet; var MasterQuery : TQuery) ; virtual ; abstract ;
    procedure SetConstQryData(sFieldName, sFieldValue:String);
//    procedure ActiveRefDataSet(RefDataSet:TDataSet);//將RefDataSet active起來==>在MstQry中Lookup到的field所屬之DataSet    
  end;

var
    dtmSTabMaintain : TdtmSTabMaintain;  
implementation

uses dtmMainU;

{$R *.DFM}

function TdtmSTabMaintain.GetMasterDataSet: TDataSet;
begin
     Result := FMstQry;
end;

procedure TdtmSTabMaintain.dtmSTabMaintainCreate(Sender: TObject);
begin
  ConstFieldNameStrList := TStringList.Create;
  ConstFieldValueStrList := TStringList.Create;
  SetMasterData(FMstQry, FMasterQuery) ;
end;


procedure TdtmSTabMaintain.SetConstQryData(sFieldName,
  sFieldValue: String);
begin
    ConstFieldNameStrList.Add(sFieldName);
    ConstFieldValueStrList.Add(sFieldValue);
end;

procedure TdtmSTabMaintain.dtmSTabMaintainDestroy(Sender: TObject);
begin
    ConstFieldNameStrList.Free;
    ConstFieldValueStrList.Free;
end;

{
procedure TdtmSTabMaintain.ActiveRefDataSet(RefDataSet: TDataSet);
begin
     RefDataSet.Close;
     RefDataSet.Open;
end;
}
function TdtmSTabMaintain.FieldValue(sI_FieldName: String): Variant;
begin
     Result := FMstQry.FieldByName(sI_FieldName).Value;
end;

procedure TdtmSTabMaintain.SingleTabDBLogin(Database: TDatabase;
  LoginParams: TStrings);
begin
//  LoginParams.Clear;
//  LoginParams.Add('UserName=sa');
//  LoginParams.Add('Password=');

end;

procedure TdtmSTabMaintain.SetMasterQuery(const Value: TQuery);
begin
  FMasterQuery := Value;
end;

function TdtmSTabMaintain.CheckDup(sI_TableName, sI_FieldName,
  sI_Value, sI_FieldType: String): Boolean;
begin
    with qryCheckDup1 do
    begin
      Close;
      SQL.Clear;
      SQL.Add('select count(*) RECCOUNT from ' + sI_TableName );
      if sI_FieldType=STR_TYPE then
        SQL.Add(' where ' + sI_FieldName + '=' + STR_SEP + sI_Value + STR_SEP )
      else if sI_FieldType=INT_TYPE then
        SQL.Add(' where ' + sI_FieldName + '='+ sI_Value);

      Open;
      if FieldByName('RECCOUNT').AsInteger>0 then
        Result := True
      else
        Result := False;
      Close;        
    end;
end;

end.
