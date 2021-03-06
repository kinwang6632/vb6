unit frmCD078B8U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ActnList, ExtCtrls, DB, ADODB, Menus,
  
  cbDBController, cbStyleController,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxImageComboBox, cxDBLookupComboBox,
  cbDataLookup, cxCurrencyEdit, cxCheckBox, cxRadioGroup, dxSkinsCore,
  dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee;

type
  TOldValue = record
    InstCode: String;
    ServiceType: String;
    GroupNo: String;
    RefNo: String;
  end;

  TfrmCD078B8 = class(TForm)
    Panel1: TPanel;
    Bevel1: TBevel;
    btnSave: TButton;
    btnCancel: TButton;
    Panel2: TPanel;
    Label1: TLabel;
    BPCode: TDataLookup;
    edtNewCodeNo: TcxTextEdit;
    radCopyNew: TcxRadioButton;
    radCopyExist: TcxRadioButton;
    CopyToBPCode: TDataLookup;
    edtNewCodeName: TcxTextEdit;
    Label2: TLabel;
    procedure BPCodeCodeNoPropertiesChange(Sender: TObject);
    procedure BPCodeCodeNamePropertiesChange(Sender: TObject);
    procedure BPCodeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure BPCodeCodeNoExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure CopyToBPCodeCodeNoExit(Sender: TObject);
    procedure CopyToBPCodeCodeNoPropertiesChange(Sender: TObject);
    procedure CopyToBPCodeCodeNamePropertiesChange(Sender: TObject);
    procedure CopyToBPCodeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure radCopyNewClick(Sender: TObject);
    procedure radCopyExistClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
  private
    FCloneCodeNo: String;
    FCloneDescription: String;
    FReader, FWriter: TADOQuery;
    FSourceLinkKeyList: TStringList;
    FCloneLinkKeyList: TStringList;
    function DataOk: Boolean;
    function GetStepNo: String;
    procedure CopyNew(TableOwner: String);
    procedure DeleteCD078X(ATableOwner: String);
    function GetCD078G1Step(const AOwner:String): String;
    { Private declarations }
  public
    { Public declarations }
    property CloneCodeNo: String read FCloneCodeNo;
    property CloneDescription: String read FCloneDescription;
  end;

var
  frmCD078B8: TfrmCD078B8;

implementation

uses cbUtilis;

{$R *.dfm}

{ TfrmCD078B2 }

procedure TfrmCD078B8.BPCodeCodeNoPropertiesChange(Sender: TObject);
begin
  BPCode.CodeNoPropertiesChange( Sender );
end;

procedure TfrmCD078B8.BPCodeCodeNamePropertiesChange(Sender: TObject);
begin
  BPCode.CodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078B8.BPCodeCodeNamePropertiesInitPopup(Sender: TObject);
begin
  BPCode.SQL.Text := Format(
      ' SELECT CODENO, DESCRIPTION ' +
      ' FROM %s.CD078              ' +
      ' WHERE STOPFLAG=0           ' +
      ' ORDER BY CODENO            ' ,
      [DBController.LoginInfo.DbAccount]);
  BPCode.CodeNamePropertiesInitPopup( Sender );
end;

procedure TfrmCD078B8.BPCodeCodeNoExit(Sender: TObject);
begin
  BPCode.CodeNoExit(Sender);
end;

procedure TfrmCD078B8.FormCreate(Sender: TObject);
begin
  BPCode.Initializa;
  CopyToBPCode.Initializa;
  FReader := DBController.DataReader;
  FWriter := DBController.DataWriter;
  FSourceLinkKeyList := TStringList.Create;
  FCloneLinkKeyList := TStringList.Create;
end;

procedure TfrmCD078B8.FormDestroy(Sender: TObject);
begin
  BPCode.Finaliza;
  CopyToBPCode.Finaliza;
  FReader.Close;
  FWriter.Close;
  FSourceLinkKeyList.Free;
  FCloneLinkKeyList.Free;
end;

procedure TfrmCD078B8.CopyToBPCodeCodeNoExit(Sender: TObject);
begin
  CopyToBPCode.CodeNoExit(Sender);
end;

procedure TfrmCD078B8.CopyToBPCodeCodeNoPropertiesChange(Sender: TObject);
begin
  CopyToBPCode.CodeNoPropertiesChange( Sender );
end;

procedure TfrmCD078B8.CopyToBPCodeCodeNamePropertiesChange(
  Sender: TObject);
begin
  CopyToBPCode.CodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078B8.CopyToBPCodeCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  CopyToBPCode.SQL.Text := Format(
      ' SELECT CODENO, DESCRIPTION ' +
      ' FROM %s.CD078              ' +
      ' WHERE STOPFLAG=0           ' +
      ' ORDER BY CODENO            ' ,
      [DBController.LoginInfo.DbAccount]);
  CopyToBPCode.CodeNamePropertiesInitPopup( Sender );
end;

procedure TfrmCD078B8.radCopyNewClick(Sender: TObject);
begin
  edtNewCodeNo.Enabled := radCopyNew.Checked;
  edtNewCodeName.Enabled := radCopyNew.Checked;
  CopyToBPCode.Enabled := not radCopyNew.Checked;
end;

procedure TfrmCD078B8.radCopyExistClick(Sender: TObject);
begin
  radCopyNewClick(Sender);
end;

procedure TfrmCD078B8.DeleteCD078X(ATableOwner: String);
var
  I:Integer;
begin
  for I := 65 to 70 do // A..F
  begin
    FWriter.Close;
    FWriter.SQL.Text := Format ( ' DELETE FROM %s.CD078%s WHERE BPCODE = ''%s'' ',
      [ATableOwner, Chr(I), FCloneCodeNo] );
    FWriter.ExecSQL;
  end;
  FWriter.Close;
  FWriter.SQL.Text := Format (
    ' DELETE FROM %s.CD078 WHERE CODENO = ''%s'' ', [ATableOwner, FCloneCodeNo] );
  FWriter.ExecSQL;
  {}
  FWriter.Close;
  FWriter.SQL.Text := Format(
    ' DELETE FROM %s.CD078A1 WHERE BPCODE = ''%s'' ', [ATableOwner, FCloneCodeNo] );
  FWriter.ExecSQL;
  {}  
  FWriter.SQL.Text := Format(
    ' DELETE FROM %s.CD078A2 WHERE BPCODE = ''%s'' ', [ATableOwner, FCloneCodeNo] );
  FWriter.ExecSQL;
  {}
  FWriter.SQL.Text := Format(
    ' DELETE FROM %s.CD078A3 WHERE BPCODE = ''%s'' ', [ATableOwner, FCloneCodeNo] );
  FWriter.ExecSQL;
  {}
  FWriter.SQL.Text := Format(
    ' DELETE FROM %s.CD078A4 WHERE BPCODE = ''%s'' ', [ATableOwner, FCloneCodeNo] );
  FWriter.ExecSQL;
  {}
  FWriter.Close;
  FWriter.SQL.Text := Format (
    ' DELETE FROM %s.CD078B1 WHERE BPCODE = ''%s'' ', [ATableOwner, FCloneCodeNo] );
  FWriter.ExecSQL;
  {}
end;

function TfrmCD078B8.GetStepNo: String;
begin
  FWriter.Close;
  FWriter.SQL.Text := Format(
    ' SELECT %s.S_CD078A1.NEXTVAL FROM DUAL ', [DBController.LoginInfo.DbAccount] );
  FWriter.Open;
  Result := FWriter.Fields[0].AsString;
  FWriter.Close;
end;

procedure TfrmCD078B8.CopyNew(TableOwner: String);
var
  aDateSt, aDateEd, aStepNo, aCD078G1Step, aLinkKey: String;
  aLinkIndex: Integer;
  aIndex: Integer;
begin
  DBController.DataConnection.BeginTrans;
  Screen.Cursor := crHourGlass;
  try
    //DeleteCD078X(TableOwner);

    //CD078
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078 WHERE CODENO=''%s'' ',
        [TableOwner, BPCode.Value] );
    FReader.Open;
    {}
    aDateSt := EmptyStr;
    aDateEd := EmptyStr;
    if ( FReader.FieldByName( 'ONSALESTARTDATE' ).AsString <> EmptyStr ) then
      aDateSt := FormatDateTime( 'yyyymmdd', FReader.FieldByName( 'ONSALESTARTDATE' ).AsDateTime );
    if ( FReader.FieldByName( 'ONSALESTOPDATE' ).AsString <> EmptyStr ) then
      aDateEd := FormatDateTime( 'yyyymmdd', FReader.FieldByName( 'ONSALESTOPDATE' ).AsDateTime );
    {}
    FWriter.Close;
    //#5232 ??????OK????,?_?????n?M??
    //#5318 ??????OK,???F?p?O????,?W?D?D??????????  By Kin 2009/10/30
    //#6267 ?W?[SYNCBPCODE???? By Kin 2012/07/19
    FWriter.SQL.Text := Format (
      ' INSERT INTO %s.CD078 (                                        ' +
      '   CODENO, DESCRIPTION,                                        ' +
      '   MASTERPROD, BUNDLE, BUNDLESAMEMON, BUNDLEMON, SAMEPERIOD,   ' +
      '   PERIOD, MERGEPRINT, PENAL, PENALTYPE, PENALSINGLE,          ' +
      '   MERGEWORK, BUNDLEPRODNOTE, WORKNOTE, REFNO, UPDTIME, UPDEN, ' +
      '   STOPFLAG, CONDITIONAL,KindFunction,ChooseKind,CanPayNow,    ' +
      '   SYNCBPCODE, CANUSETYPE, EPGIMAGEURL, EPGORD )               ' +
      ' VALUES (                                                      ' +
      '   ''%s'', ''%s'',                                             ' +
      '   ''%s'', ''%s'',                                             ' +
      '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',     ' +
      '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',     ' +
      '   ''%s'', ''%s'',''%s'',''%s'',''%s'',''%s'',                 ' +
      '   ''%s'', ''%s'',''%s''  )                                    ',
      [TableOwner,
      FCloneCodeNo,
      FCloneDescription,
      FReader.FieldByName('MASTERPROD').AsString,
      FReader.FieldByName('BUNDLE').AsString,
      FReader.FieldByName('BUNDLESAMEMON').AsString,
      FReader.FieldByName('BUNDLEMON').AsString,
      FReader.FieldByName('SAMEPERIOD').AsString,
      FReader.FieldByName('PERIOD').AsString,
      FReader.FieldByName('MERGEPRINT').AsString,
      FReader.FieldByName('PENAL').AsString,
      FReader.FieldByName('PENALTYPE').AsString,
      FReader.FieldByName('PENALSINGLE').AsString,
      FReader.FieldByName('MERGEWORK').AsString,
      FReader.FieldByName('BUNDLEPRODNOTE').AsString,
      FReader.FieldByName('WORKNOTE').AsString,
      FReader.FieldByName('REFNO').AsString,
      FReader.FieldByName('UPDTIME').AsString,
      FReader.FieldByName('UPDEN').AsString,
      FReader.FieldByName('STOPFLAG').AsString,
      FReader.FieldByName('CONDITIONAL').AsString,
      FReader.FieldByName('KindFunction').AsString,
      FReader.FieldByName('ChooseKind').AsString,
      FReader.FieldByName('CANPAYNOW').AsString,
      FReader.FieldByName('SYNCBPCODE').AsString,
      FReader.FieldByName('CANUSETYPE').AsString,
      FReader.FieldByName('EPGIMAGEURL').AsString,
      FReader.FieldByName('EPGORD').AsString] );
    FWriter.ExecSQL;

    //CD078A
    //#5318 ??????OK ???s?L?hCD078A???F?I?????k?N?X???W???? By Kin 2009/10/30
    //#5913 ?W?[PFlag2 By Kin 2011/04/14
    //7130 There was lost to update AuthTunerCount field of value By Kin 2015/11/09
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078A WHERE BPCODE=''%s'' ',
        [TableOwner, BPCode.Value] );
    FReader.Open;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078A ( bpcode, citemcode, citemname, servicetype,' +
        '   faciitem, bundle, bundlemon, penaltype, expiretype, expirecut,  ' +
        '   depositattr, depositcode, depositname, monthamt,                ' +
        '   dayamt, pflag1, pflagdate, truncamt, pflagcode, pflagname,      ' +
        '   mon1, period1, ratetype1, discountamt1, monthamt1, dayamt1,     ' +
        '   punish1, mon2, period2, ratetype2, discountamt2, monthamt2,     ' +
        '   dayamt2, punish2, mon3, period3, ratetype3, discountamt3,       ' +
        '   monthamt3, dayamt3, punish3, mon4, period4, ratetype4,          ' +
        '   discountamt4, monthamt4, dayamt4, punish4, mon5, period5,       ' +
        '   ratetype5, discountamt5, monthamt5, dayamt5, punish5, mon6,     ' +
        '   period6, ratetype6, discountamt6, monthamt6, dayamt6, punish6,  ' +
        '   mon7, period7, ratetype7, discountamt7, monthamt7, dayamt7,     ' +
        '   punish7, mon8, period8, ratetype8, discountamt8, monthamt8,     ' +
        '   dayamt8, punish8, mon9, period9, ratetype9, discountamt9,       ' +
        '   monthamt9, dayamt9, punish9, mon10, period10, ratetype10,       ' +
        '   discountamt10, monthamt10, dayamt10, punish10, mon11, period11, ' +
        '   ratetype11, discountamt11, monthamt11, dayamt11, punish11,      ' +
        '   mon12, period12, ratetype12, discountamt12, monthamt12,         ' +
        '   dayamt12, punish12, penal, penalcode, penalname,                ' +
        '   depositcodestr, discountcalc, mcitemcode,                       ' +
        '   SalePointcode,SalePointName,PFlag2,AuthTunerCount )             ' +
        ' VALUES ( ''%s'', ''%s'', ''%s'', ''%s'',                          ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'',                                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '   ''%s'', ''%s'', ''%s'',''%s'',''%s'',''%s'',''%s'' )            ',
        [TableOwner,
         FCloneCodeNo,
         FReader.FieldByName( 'citemcode' ).AsString,
         FReader.FieldByName( 'citemname' ).AsString,
         FReader.FieldByName( 'servicetype' ).AsString,
         FReader.FieldByName( 'faciitem' ).AsString,
         FReader.FieldByName( 'bundle' ).AsString,
         FReader.FieldByName( 'bundlemon' ).AsString,
         FReader.FieldByName( 'penaltype' ).AsString,
         FReader.FieldByName( 'expiretype' ).AsString,
         FReader.FieldByName( 'expirecut' ).AsString,
         FReader.FieldByName( 'depositattr' ).AsString,
         FReader.FieldByName( 'depositcode' ).AsString,
         FReader.FieldByName( 'depositname' ).AsString,
         FReader.FieldByName( 'monthamt' ).AsString,
         FReader.FieldByName( 'dayamt' ).AsString,
         FReader.FieldByName( 'pflag1' ).AsString,
         FReader.FieldByName( 'pflagdate' ).AsString,
         FReader.FieldByName( 'truncamt' ).AsString,
         FReader.FieldByName( 'pflagcode' ).AsString,
         FReader.FieldByName( 'pflagname' ).AsString,
         {}
         FReader.FieldByName( 'mon1').AsString,
         FReader.FieldByName( 'period1').AsString,
         FReader.FieldByName( 'ratetype1' ).AsString,
         FReader.FieldByName( 'discountamt1' ).AsString,
         FReader.FieldByName( 'monthamt1' ).AsString,
         FReader.FieldByName( 'dayamt1' ).AsString,
         FReader.FieldByName( 'punish1' ).AsString,
         {}
         FReader.FieldByName( 'mon2').AsString,
         FReader.FieldByName( 'period2').AsString,
         FReader.FieldByName( 'ratetype2' ).AsString,
         FReader.FieldByName( 'discountamt2' ).AsString,
         FReader.FieldByName( 'monthamt2' ).AsString,
         FReader.FieldByName( 'dayamt2' ).AsString,
         FReader.FieldByName( 'punish2' ).AsString,
         {}
         FReader.FieldByName( 'mon3').AsString,
         FReader.FieldByName( 'period3').AsString,
         FReader.FieldByName( 'ratetype3' ).AsString,
         FReader.FieldByName( 'discountamt3' ).AsString,
         FReader.FieldByName( 'monthamt3' ).AsString,
         FReader.FieldByName( 'dayamt3' ).AsString,
         FReader.FieldByName( 'punish3' ).AsString,
         {}
         FReader.FieldByName( 'mon4').AsString,
         FReader.FieldByName( 'period4').AsString,
         FReader.FieldByName( 'ratetype4' ).AsString,
         FReader.FieldByName( 'discountamt4' ).AsString,
         FReader.FieldByName( 'monthamt4' ).AsString,
         FReader.FieldByName( 'dayamt4' ).AsString,
         FReader.FieldByName( 'punish4' ).AsString,
         {}
         FReader.FieldByName( 'mon5').AsString,
         FReader.FieldByName( 'period5').AsString,
         FReader.FieldByName( 'ratetype5' ).AsString,
         FReader.FieldByName( 'discountamt5' ).AsString,
         FReader.FieldByName( 'monthamt5' ).AsString,
         FReader.FieldByName( 'dayamt5' ).AsString,
         Freader.FieldByName( 'punish5' ).AsString,
         {}
         FReader.FieldByName( 'mon6' ).AsString,
         FReader.FieldByName( 'period6' ).AsString,
         FReader.FieldByName( 'ratetype6' ).AsString,
         FReader.FieldByName( 'discountamt6' ).AsString,
         FReader.FieldByName( 'monthamt6' ).AsString,
         FReader.FieldByName( 'dayamt6' ).AsString,
         FReader.FieldByName( 'punish6' ).AsString,
         {}
         FReader.FieldByName( 'mon7').AsString,
         FReader.FieldByName( 'period7').AsString,
         FReader.FieldByName( 'ratetype7' ).AsString,
         FReader.FieldByName( 'discountamt7' ).AsString,
         FReader.FieldByName( 'monthamt7' ).AsString,
         FReader.FieldByName( 'dayamt7' ).AsString,
         FReader.FieldByName( 'punish7' ).AsString,
         {}
         FReader.FieldByName( 'mon8').AsString,
         FReader.FieldByName( 'period8').AsString,
         FReader.FieldByName( 'ratetype8' ).AsString,
         FReader.FieldByName( 'discountamt8' ).AsString,
         FReader.FieldByName( 'monthamt8' ).AsString,
         FReader.FieldByName( 'dayamt8' ).AsString,
         FReader.FieldByName( 'punish8' ).AsString,
         {}
         FReader.FieldByName( 'mon9').AsString,
         FReader.FieldByName( 'period9').AsString,
         FReader.FieldByName( 'ratetype9' ).AsString,
         FReader.FieldByName( 'discountamt9' ).AsString,
         FReader.FieldByName( 'monthamt9' ).AsString,
         FReader.FieldByName( 'dayamt9' ).AsString,
         FReader.FieldByName( 'punish9' ).AsString,
         {}
         FReader.FieldByName( 'mon10').AsString,
         FReader.FieldByName( 'period10').AsString,
         FReader.FieldByName( 'ratetype10' ).AsString,
         FReader.FieldByName( 'discountamt10' ).AsString,
         FReader.FieldByName( 'monthamt10' ).AsString,
         FReader.FieldByName( 'dayamt10' ).AsString,
         FReader.FieldByName( 'punish10' ).AsString,
         {}
         FReader.FieldByName( 'mon11').AsString,
         FReader.FieldByName( 'period11').AsString,
         FReader.FieldByName( 'ratetype11' ).AsString,
         FReader.FieldByName( 'discountamt11' ).AsString,
         FReader.FieldByName( 'monthamt11' ).AsString,
         FReader.FieldByName( 'dayamt11' ).AsString,
         FReader.FieldByName( 'punish11' ).AsString,
         {}
         FReader.FieldByName( 'mon12').AsString,
         FReader.FieldByName( 'period12').AsString,
         FReader.FieldByName( 'ratetype12' ).AsString,
         FReader.FieldByName( 'discountamt12' ).AsString,
         FReader.FieldByName( 'monthamt12' ).AsString,
         FReader.FieldByName( 'dayamt12' ).AsString,
         FReader.FieldByName( 'punish12' ).AsString,
         {}
         FReader.FieldByName( 'penal' ).AsString,
         FReader.FieldByName( 'penalcode' ).AsString,
         FReader.FieldByName( 'penalname' ).AsString,
         {}
         FReader.FieldByName( 'depositcodestr' ).AsString,
         FReader.FieldByName( 'discountcalc' ).AsString,
         FReader.FieldByName( 'mcitemcode' ).AsString,
         FReader.FieldByName( 'SalePointcode').AsString,
         FReader.FieldByName( 'SalePointName').AsString,
         FReader.FieldByName( 'PFlag2' ).AsString,
         FReader.FieldByName( 'AuthTunerCount' ).AsString ] );
      FWriter.ExecSQL;
      {}
      FWriter.SQL.Text := EmptyStr;
      FWriter.Parameters.Clear;
      {}
      {#5257 ?W?[?S?w?u?f?I???? By Kin 2010/05/05}
      FWriter.SQL.Text := Format(
        ' UPDATE %s.CD078A ' +
        '    SET INSTCODESTR = :INSTCODESTR,  ' +
        '        COMMENT1 = :COMMENT1,         ' +
        '        COMMENT2 = :COMMENT2,         ' +
        '        COMMENT3 = :COMMENT3,         ' +
        '        COMMENT4 = :COMMENT4,         ' +
        '        COMMENT5 = :COMMENT5,         ' +
        '        COMMENT6 = :COMMENT6,         ' +
        '        COMMENT7 = :COMMENT7,         ' +
        '        COMMENT8 = :COMMENT8,         ' +
        '        COMMENT9 = :COMMENT9,         ' +
        '        COMMENT10 = :COMMENT10,       ' +
        '        COMMENT11 = :COMMENT11,       ' +
        '        COMMENT12 = :COMMENT12,       ' +
        '        CMBAUDNO = :CMBAUDNO,         ' +
        '        CMBAUDRATE = :CMBAUDRATE,     ' +
        '        BPSTOPDATE = :BPSTOPDATE      ' +
        '  WHERE BPCODE = ''%s''               ' +
        '    AND CITEMCODE = ''%s''            ' +
        '    AND SERVICETYPE = ''%s''          ',
        [TableOwner, FCloneCodeNo,
         FReader.FieldByName( 'CITEMCODE' ).AsString,
         FReader.FieldByName( 'SERVICETYPE' ).AsString] );
      {}   
      FWriter.Parameters.ParamByName( 'INSTCODESTR' ).Value := FReader.FieldByName( 'INSTCODESTR' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT1' ).Value := FReader.FieldByName( 'COMMENT1' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT2' ).Value := FReader.FieldByName( 'COMMENT2' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT3' ).Value := FReader.FieldByName( 'COMMENT3' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT4' ).Value := FReader.FieldByName( 'COMMENT4' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT5' ).Value := FReader.FieldByName( 'COMMENT5' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT6' ).Value := FReader.FieldByName( 'COMMENT6' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT7' ).Value := FReader.FieldByName( 'COMMENT7' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT8' ).Value := FReader.FieldByName( 'COMMENT8' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT9' ).Value := FReader.FieldByName( 'COMMENT9' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT10' ).Value := FReader.FieldByName( 'COMMENT10' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT11' ).Value := FReader.FieldByName( 'COMMENT11' ).AsString;
      FWriter.Parameters.ParamByName( 'COMMENT12' ).Value := FReader.FieldByName( 'COMMENT12' ).AsString;
      FWriter.Parameters.ParamByName( 'CMBAUDNO' ).Value := FReader.FieldByName( 'CMBAUDNO' ).AsString;
      FWriter.Parameters.ParamByName( 'CMBAUDRATE' ).Value := FReader.FieldByName( 'CMBAUDRATE' ).AsString;
      if FReader.FieldByName( 'BPSTOPDATE' ).AsString = EmptyStr then
        FWriter.Parameters.ParamByName( 'BPSTOPDATE' ).Value := EmptyStr
      else
        FWriter.Parameters.ParamByName( 'BPSTOPDATE' ).Value := FReader.FieldByName( 'BPSTOPDATE' ).AsDateTime;
      FWriter.ExecSQL;
      FReader.Next
    end;

    FSourceLinkKeyList.Clear;
    FCloneLinkKeyList.Clear;

    //CD078A1, PK ?O StepNo, Copy ?\???n?b???@?? StepNo, ???M?| PK????
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078A1  ' +
        '  WHERE BPCODE=''%s''      ' +
        '  ORDER BY BPCODE, CITEMCODE, SERVICETYPE, FACIITEM, STEPNO ',
        [TableOwner, BPCode.Value] );
    FReader.Open;
    {#6945 Period ?S???????? By Kin 2014/11/28}
    while not FReader.Eof do
    begin
      FSourceLinkKeyList.Add( FReader.FieldByName( 'STEPNO' ).AsString );
      {}
      aStepNo := GetStepNo;
      {}
      FCloneLinkKeyList.Add( aStepNo );
      {}
      aLinkKey := EmptyStr;
      aLinkIndex := -1;

      if ( FReader.FieldByName( 'LINKKEY' ).AsString <> EmptyStr ) then
        aLinkIndex := FSourceLinkKeyList.IndexOf( FReader.FieldByName( 'LINKKEY' ).AsString );
      if ( aLinkIndex >= 0 ) then
        aLinkKey := FCloneLinkKeyList[aLinkIndex];
      {}  
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078A1( bpcode, citemcode, servicetype,           ' +
        '   faciitem, StepNo, Description, MasterSale,                      ' +
        '   mon1, period1, ratetype1, discountamt1, monthamt1, dayamt1,     ' +
        '   punish1, mon2, period2, ratetype2, discountamt2, monthamt2,     ' +
        '   dayamt2, punish2, mon3, period3, ratetype3, discountamt3,       ' +
        '   monthamt3, dayamt3, punish3, mon4, period4, ratetype4,          ' +
        '   discountamt4, monthamt4, dayamt4, punish4, mon5, period5,       ' +
        '   ratetype5, discountamt5, monthamt5, dayamt5, punish5, mon6,     ' +
        '   period6, ratetype6, discountamt6, monthamt6, dayamt6, punish6,  ' +
        '   mon7, period7, ratetype7, discountamt7, monthamt7, dayamt7,     ' +
        '   punish7, mon8, period8, ratetype8, discountamt8, monthamt8,     ' +
        '   dayamt8, punish8, mon9, period9, ratetype9, discountamt9,       ' +
        '   monthamt9, dayamt9, punish9, mon10, period10, ratetype10,       ' +
        '   discountamt10, monthamt10, dayamt10, punish10, mon11, period11, ' +
        '   ratetype11, discountamt11, monthamt11, dayamt11, punish11,      ' +
        '   mon12, period12, ratetype12, discountamt12, monthamt12,         ' +
        '   dayamt12, punish12, linkkey, Period, StopFlag )                 ' +
        ' VALUES ( ''%s'', ''%s'', ''%s'',                                  ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'',                                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '   ''%s'', ''%s'', ''%s'',''%s'', ''%s'' )                         ',
        [TableOwner,
         FCloneCodeNo,
         FReader.FieldByName( 'citemcode' ).AsString,
         FReader.FieldByName( 'servicetype' ).AsString,
         FReader.FieldByName( 'faciItem' ).AsString,
         aStepNo,
         FReader.FieldByName( 'description' ).AsString,
         FReader.FieldByName( 'mastersale' ).AsString,
         {}
         FReader.FieldByName( 'mon1').AsString,
         FReader.FieldByName( 'period1').AsString,
         FReader.FieldByName( 'ratetype1' ).AsString,
         FReader.FieldByName( 'discountamt1' ).AsString,
         FReader.FieldByName( 'monthamt1' ).AsString,
         FReader.FieldByName( 'dayamt1' ).AsString,
         FReader.FieldByName( 'punish1' ).AsString,
         {}
         FReader.FieldByName( 'mon2').AsString,
         FReader.FieldByName( 'period2').AsString,
         FReader.FieldByName( 'ratetype2' ).AsString,
         FReader.FieldByName( 'discountamt2' ).AsString,
         FReader.FieldByName( 'monthamt2' ).AsString,
         FReader.FieldByName( 'dayamt2' ).AsString,
         FReader.FieldByName( 'punish2' ).AsString,
         {}
         FReader.FieldByName( 'mon3').AsString,
         FReader.FieldByName( 'period3').AsString,
         FReader.FieldByName( 'ratetype3' ).AsString,
         FReader.FieldByName( 'discountamt3' ).AsString,
         FReader.FieldByName( 'monthamt3' ).AsString,
         FReader.FieldByName( 'dayamt3' ).AsString,
         FReader.FieldByName( 'punish3' ).AsString,
         {}
         FReader.FieldByName( 'mon4').AsString,
         FReader.FieldByName( 'period4').AsString,
         FReader.FieldByName( 'ratetype4' ).AsString,
         FReader.FieldByName( 'discountamt4' ).AsString,
         FReader.FieldByName( 'monthamt4' ).AsString,
         FReader.FieldByName( 'dayamt4' ).AsString,
         FReader.FieldByName( 'punish4' ).AsString,
         {}
         FReader.FieldByName( 'mon5').AsString,
         FReader.FieldByName( 'period5').AsString,
         FReader.FieldByName( 'ratetype5' ).AsString,
         FReader.FieldByName( 'discountamt5' ).AsString,
         FReader.FieldByName( 'monthamt5' ).AsString,
         FReader.FieldByName( 'dayamt5' ).AsString,
         FReader.FieldByName( 'punish5' ).AsString,
         {}
         FReader.FieldByName( 'mon6').AsString,
         FReader.FieldByName( 'period6').AsString,
         FReader.FieldByName( 'ratetype6' ).AsString,
         FReader.FieldByName( 'discountamt6' ).AsString,
         FReader.FieldByName( 'monthamt6' ).AsString,
         FReader.FieldByName( 'dayamt6' ).AsString,
         FReader.FieldByName( 'punish6' ).AsString,
         {}
         FReader.FieldByName( 'mon7').AsString,
         FReader.FieldByName( 'period7').AsString,
         FReader.FieldByName( 'ratetype7' ).AsString,
         FReader.FieldByName( 'discountamt7' ).AsString,
         FReader.FieldByName( 'monthamt7' ).AsString,
         FReader.FieldByName( 'dayamt7' ).AsString,
         FReader.FieldByName( 'punish7' ).AsString,
         {}
         FReader.FieldByName( 'mon8').AsString,
         FReader.FieldByName( 'period8').AsString,
         FReader.FieldByName( 'ratetype8' ).AsString,
         FReader.FieldByName( 'discountamt8' ).AsString,
         FReader.FieldByName( 'monthamt8' ).AsString,
         FReader.FieldByName( 'dayamt8' ).AsString,
         FReader.FieldByName( 'punish8' ).AsString,
         {}
         FReader.FieldByName( 'mon9').AsString,
         FReader.FieldByName( 'period9').AsString,
         FReader.FieldByName( 'ratetype9' ).AsString,
         FReader.FieldByName( 'discountamt9' ).AsString,
         FReader.FieldByName( 'monthamt9' ).AsString,
         FReader.FieldByName( 'dayamt9' ).AsString,
         FReader.FieldByName( 'punish9' ).AsString,
         {}
         FReader.FieldByName( 'mon10').AsString,
         FReader.FieldByName( 'period10').AsString,
         FReader.FieldByName( 'ratetype10' ).AsString,
         FReader.FieldByName( 'discountamt10' ).AsString,
         FReader.FieldByName( 'monthamt10' ).AsString,
         FReader.FieldByName( 'dayamt10' ).AsString,
         FReader.FieldByName( 'punish10' ).AsString,
         {}
         FReader.FieldByName( 'mon11').AsString,
         FReader.FieldByName( 'period11').AsString,
         FReader.FieldByName( 'ratetype11' ).AsString,
         FReader.FieldByName( 'discountamt11' ).AsString,
         FReader.FieldByName( 'monthamt11' ).AsString,
         FReader.FieldByName( 'dayamt11' ).AsString,
         FReader.FieldByName( 'punish11' ).AsString,
         {}
         FReader.FieldByName( 'mon12').AsString,
         FReader.FieldByName( 'period12').AsString,
         FReader.FieldByName( 'ratetype12' ).AsString,
         FReader.FieldByName( 'discountamt12' ).AsString,
         FReader.FieldByName( 'monthamt12' ).AsString,
         FReader.FieldByName( 'dayamt12' ).AsString,
         FReader.FieldByName( 'punish12' ).AsString,
         aLinkKey,
         FReader.FieldByName('Period').AsString,
         FReader.FieldByName('StopFlag').AsString] );
      FWriter.ExecSQL;
      {}
      FWriter.SQL.Text := EmptyStr;
      FWriter.Parameters.Clear;
      {}
      FWriter.SQL.Text := Format(
        ' UPDATE %s.CD078A1 ' +
        '    SET COMMENT1 = :COMMENT1,         ' +
        '        COMMENT2 = :COMMENT2,         ' +
        '        COMMENT3 = :COMMENT3,         ' +
        '        COMMENT4 = :COMMENT4,         ' +
        '        COMMENT5 = :COMMENT5,         ' +
        '        COMMENT6 = :COMMENT6,         ' +
        '        COMMENT7 = :COMMENT7,         ' +
        '        COMMENT8 = :COMMENT8,         ' +
        '        COMMENT9 = :COMMENT9,         ' +
        '        COMMENT10 = :COMMENT10,       ' +
        '        COMMENT11 = :COMMENT11,       ' +
        '        COMMENT12 = :COMMENT12,       ' +
        '        DISCOUNTRATE1 = :DISCOUNTRATE1,   ' +
        '        DISCOUNTRATE2 = :DISCOUNTRATE2,   ' +
        '        DISCOUNTRATE3 = :DISCOUNTRATE3,   ' +
        '        DISCOUNTRATE4 = :DISCOUNTRATE4,   ' +
        '        DISCOUNTRATE5 = :DISCOUNTRATE5,   ' +
        '        DISCOUNTRATE6 = :DISCOUNTRATE6,   ' +
        '        DISCOUNTRATE7 = :DISCOUNTRATE7,   ' +
        '        DISCOUNTRATE8 = :DISCOUNTRATE8,   ' +
        '        DISCOUNTRATE9 = :DISCOUNTRATE9,   ' +
        '        DISCOUNTRATE10 = :DISCOUNTRATE10, ' +
        '        DISCOUNTRATE11 = :DISCOUNTRATE11, ' +
        '        DISCOUNTRATE12 = :DISCOUNTRATE12  ' +
        '  WHERE BPCODE = ''%s''               ' +
        '    AND STEPNO = ''%s''               ',
        [TableOwner, FCloneCodeNo, aStepNo] );
      for aIndex := 1 to 12 do
      begin
        FWriter.Parameters.ParamByName( Format( 'COMMENT%d', [aIndex] ) ).Value :=
          FReader.FieldByName( Format( 'COMMENT%d', [aIndex] ) ).AsString;
        {}  
        FWriter.Parameters.ParamByName( Format( 'DISCOUNTRATE%d', [aIndex] ) ).Value :=
          FReader.FieldByName( Format( 'DISCOUNTRATE%d', [aIndex] ) ).AsString;
      end;
      FWriter.ExecSQL;
      {}
      FReader.Next;
    end;

    FSourceLinkKeyList.Clear;
    FCloneLinkKeyList.Clear;

    
    //CD078A2
    FReader.Close;
    FReader.SQL.Text := Format(
      ' SELECT * FROM %s.CD078A2 WHERE BPCODE = ''%s'' ', [TableOwner, BPCode.Value] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format(
        ' INSERT INTO %s.CD078A2 (                               ' +
        '   BPCODE, CITEMCODE, SERVICETYPE, DEPOSITCODE,         ' +
        '   PTCODE, PTNAME, DEPOSITAMT, DEPOSITDEFAULT  )        ' +
        ' VALUES (                                               ' +
        '  ''%s'', ''%s'', ''%s'', ''%s'',                       ' +
        '  ''%s'', ''%s'', ''%s'', ''%s''  )                     ',
        [TableOwner,
         FCloneCodeNo,
         FReader.FieldByName( 'CITEMCODE' ).AsString,
         FReader.FieldByName( 'SERVICETYPE' ).AsString,
         FReader.FieldByName( 'DEPOSITCODE' ).AsString,
         FReader.FieldByName( 'PTCODE' ).AsString,
         FReader.FieldByName( 'PTNAME' ).AsString,
         FReader.FieldByName( 'DEPOSITAMT' ).AsString,
         FReader.FieldByName( 'DEPOSITDEFAULT' ).AsString] );
      FWriter.ExecSQL;   
      FReader.Next;
    end;


    //CD078A3
    FReader.Close;
    FReader.SQL.Text := Format(
      ' SELECT * FROM %s.CD078A3 WHERE BPCODE = ''%s'' ', [TableOwner, BPCode.Value] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format(
        ' INSERT INTO %s.CD078A3 (                               ' +
        '   BPCODE, CITEMCODE, SERVICETYPE, PENALCODE,            ' +
        '   ITEM, PENALMON, PENALAMT  )                          ' +
        ' VALUES (                                               ' +
        '  ''%s'', ''%s'', ''%s'', ''%s'',                       ' +
        '  ''%s'', ''%s'', ''%s''  )                             ',
        [TableOwner,
         FCloneCodeNo,
         FReader.FieldByName( 'CITEMCODE' ).AsString,
         FReader.FieldByName( 'SERVICETYPE' ).AsString,
         FReader.FieldByName( 'PENALCODE' ).AsString,
         FReader.FieldByName( 'ITEM' ).AsString,
         FReader.FieldByName( 'PENALMON' ).AsString,
         FReader.FieldByName( 'PENALAMT' ).AsString] );
      FWriter.ExecSQL;   
      FReader.Next;
    end;

    //CD078A4
    FReader.Close;
    FReader.SQL.Text := Format(
      ' SELECT * FROM %s.CD078A4 WHERE BPCODE = ''%s'' ', [TableOwner, BPCode.Value] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format(
        ' INSERT INTO %s.CD078A4 (                                 ' +
        '   BPCODE, CITEMCODE, SERVICETYPE, PENALCODE,             ' +
        '   ITEM, MONTHSTART, MONTHSTOP, MONTHAMT, DECREASEAMT )   ' +
        ' VALUES (                                                 ' +
        '  ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
        '  ''%s'', ''%s'', ''%s'', ''%s'', ''%s''  )               ',
        [TableOwner,
         FCloneCodeNo,
         FReader.FieldByName( 'CITEMCODE' ).AsString,
         FReader.FieldByName( 'SERVICETYPE' ).AsString,
         FReader.FieldByName( 'PENALCODE' ).AsString,
         FReader.FieldByName( 'ITEM' ).AsString,
         FReader.FieldByName( 'MONTHSTART' ).AsString,
         FReader.FieldByName( 'MONTHSTOP' ).AsString,
         FReader.FieldByName( 'MONTHAMT' ).AsString,
         FReader.FieldByName( 'DECREASEAMT' ).AsString] );
      FWriter.ExecSQL;   
      FReader.Next;
    end;

    //CD078B
    FReader.Close;
    FReader.SQL.Text := Format (
      ' SELECT * FROM %s.CD078B WHERE BPCODE=''%s'' ', [TableOwner, BPCode.Value] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078B (                                        ' +
        '   BPCODE, NEWBPCODE, INSTCODE, INSTNAME, SERVICETYPE, GROUPNO, ' +
        '   GROUPNAME, WORKUNIT, REFNO, INTERDEPEND,                     ' +
        '   PRCODE, PRNAME  )                                            ' +
        ' VALUES (                                                       ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',      ' +
        '   ''%s'', ''%s'', ''%s'',                                      ' +
        '   ''%s'', ''%s''  )                                            ',
        [TableOwner,
        FCloneCodeNo,
        FReader.FieldByName('NEWBPCODE').AsString,
        FReader.FieldByName('INSTCODE').AsString,
        FReader.FieldByName('INSTNAME').AsString,
        FReader.FieldByName('SERVICETYPE').AsString,
        FReader.FieldByName('GROUPNO').AsString,
        FReader.FieldByName('GROUPNAME').AsString,
        FReader.FieldByName('WORKUNIT').AsString,
        FReader.FieldByName('REFNO').AsString,
        FReader.FieldByName('INTERDEPEND').AsString,
        FReader.FieldByName('PRCODE').AsString,
        FReader.FieldByName('PRNAME').AsString] );
      FWriter.ExecSQL;
      FReader.Next
    end;


    //CD078B1
    FReader.Close;
    FReader.SQL.Text := Format (
      ' SELECT * FROM %s.CD078B1 WHERE BPCODE=''%s'' ', [TableOwner, BPCode.Value] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078B1 (                                       ' +
        '   BPCODE, INSTCODE, SERVICETYPE, INSTCODE2, INSTNAME2,         ' +
        '   PRCODE2, PRNAME2  )                                          ' +
        ' VALUES (                                                       ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s''  )    ',
        [TableOwner,
         FCloneCodeNo,
         FReader.FieldByName( 'INSTCODE' ).AsString,
         FReader.FieldByName( 'SERVICETYPE' ).AsString,
         FReader.FieldByName( 'INSTCODE2' ).AsString,
         FReader.FieldByName( 'INSTNAME2' ).AsString,
         FReader.FieldByName( 'PRCODE2' ).AsString,
         FReader.FieldByName( 'PRNAME2' ).AsString] );
      FWriter.ExecSQL;
      FReader.Next
    end;


    //CD078C
    FReader.Close;
    FReader.SQL.Text := Format (
      ' SELECT * FROM %s.CD078C WHERE BPCODE=''%s'' ', [TableOwner, BPCode.Value] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078C (                                        ' +
        '   BPCODE, NEWBPCODE, CITEMCODE, CITEMNAME, SERVICETYPE,        ' +
        '   FACIITEM, AMOUNT, DISCOUNTAMT, PUNISH, PENALTYPE,            ' +
        '   REFCITEMCODE, REFCITEMNAME, INSTCODESTR )                    ' +
        ' VALUES (                                                       ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',      ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'' )             ',
        [TableOwner,
        FCloneCodeNo,
        FReader.FieldByName('NEWBPCODE').AsString,
        FReader.FieldByName('CITEMCODE').AsString,
        FReader.FieldByName('CITEMNAME').AsString,
        FReader.FieldByName('SERVICETYPE').AsString,
        FReader.FieldByName('FACIITEM').AsString,
        FReader.FieldByName('AMOUNT').AsString,
        FReader.FieldByName('DISCOUNTAMT').AsString,
        FReader.FieldByName('PUNISH').AsString,
        FReader.FieldByName('PENALTYPE').AsString,
        FReader.FieldByName('REFCITEMCODE').AsString,
        FReader.FieldByName('REFCITEMNAME').AsString,
        FReader.FieldByName('INSTCODESTR').AsString] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078D
    FReader.Close;
    FReader.SQL.Text := Format (
      ' SELECT * FROM %s.CD078D WHERE BPCODE=''%s'' ', [TableOwner, BPCode.Value] );
    FReader.Open;
    FReader.First;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078D (                                        ' +
        '   BPCODE, NEWBPCODE, FACIITEM, FACITYPE, FACICODE, FACINAME,   ' +
        '   SERVICETYPE, MODELCODE, MODELNAME, BUYCODE, BUYNAME,         ' +
        '   CMBAUDNO, CMBAUDRATE, FIXIPCOUNT, DYNIPCOUNT, EMCCM,         ' +
        '   EMCCMCP, CMCPDYNIP, CMCPFIXIP, INSTCODESTR,                  ' +
        '   CBPCHANGEFIXIP, CBPCHANGEDYNIP )                             ' +
        ' VALUES (                                                       ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',      ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',      ' +
        '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',              ' +
        '   ''%s'', ''%s''  )                                            ',
        [TableOwner,
        FCloneCodeNo,
        FReader.FieldByName('NEWBPCODE').AsString,
        FReader.FieldByName('FACIITEM').AsString,
        FReader.FieldByName('FACITYPE').AsString,
        FReader.FieldByName('FACICODE').AsString,
        FReader.FieldByName('FACINAME').AsString,
        FReader.FieldByName('SERVICETYPE').AsString,
        FReader.FieldByName('MODELCODE').AsString,
        FReader.FieldByName('MODELNAME').AsString,
        FReader.FieldByName('BUYCODE').AsString,
        FReader.FieldByName('BUYNAME').AsString,
        FReader.FieldByName('CMBAUDNO').AsString,
        FReader.FieldByName('CMBAUDRATE').AsString,
        FReader.FieldByName('FIXIPCOUNT').AsString,
        FReader.FieldByName('DYNIPCOUNT').AsString,
        FReader.FieldByName('EMCCM').AsString,
        FReader.FieldByName('EMCCMCP').AsString,
        FReader.FieldByName('CMCPDYNIP').AsString,
        FReader.FieldByName('CMCPFIXIP').AsString,
        FReader.FieldByName('INSTCODESTR').AsString,
        FReader.FieldByName( 'CBPCHANGEFIXIP' ).AsString,
        FReader.FieldByName( 'CBPCHANGEDYNIP' ).AsString] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078E
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078E WHERE BPCODE=''%s'' ', [TableOwner, BPCode.Value] );
    FReader.Open;
    while not FReader.Eof do
    begin
      {}
      aDateSt := EmptyStr;
      aDateEd := EmptyStr;
      if ( FReader.FieldByName( 'ONSALESTARTDATE' ).AsString <> EmptyStr ) then
        aDateSt := FormatDateTime( 'yyyymmdd', FReader.FieldByName( 'ONSALESTARTDATE' ).AsDateTime );
      if ( FReader.FieldByName( 'ONSALESTOPDATE' ).AsString <> EmptyStr ) then
        aDateEd := FormatDateTime( 'yyyymmdd', FReader.FieldByName( 'ONSALESTOPDATE' ).AsDateTime );
      FWriter.Close;
      FWriter.SQL.Text := Format (
          ' INSERT INTO %s.CD078E (                                               ' +
          '   BPCODE, NEWBPCODE, NEWBPNAME, CHANGESTARTDATE, CHANGESTOPDATE )     ' +
          ' VALUES (                                                              ' +
          '   ''%s'', ''%s'', ''%s'', TO_DATE( ''%s'', ''YYYYMMDD'' ),            ' +
          '   TO_DATE( ''%s'', ''YYYYMMDD'' )                                     ',
          [TableOwner,
          FCloneCodeNo,
          FReader.FieldByName('NEWBPCODE').AsString,
          FReader.FieldByName('NEWBPNAME').AsString,
          aDateSt,
          aDateEd] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078F
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078F WHERE BPCODE=''%s'' ',
        [TableOwner, BPCode.Value] );
    FReader.Open;
    while not FReader.Eof do
      begin
        FWriter.Close;
        FWriter.SQL.Text := Format (
            ' INSERT INTO %s.CD078F (                                               ' +
            '   BPCODE, NEWBPCODE, SERVICETYPE, OLDCITEMCODE, OLDCITEMNAME,         ' +
            '   NEWCITEMCODE, NEWCITEMNAME, CHANGETYPE, PENALTYPE, PENAL,           ' +
            '   DEPOSITTYPE, FACIITEM )                                             ' +
            ' VALUES (                                                              ' +
            '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',     ' +
            '   ''%s'', ''%s'', ''%s'', ''%s'' )                                    ',
            [TableOwner,
            FCloneCodeNo,
            FReader.FieldByName('NEWBPCODE').AsString,
            FReader.FieldByName('SERVICETYPE').AsString,
            FReader.FieldByName('OLDCITEMCODE').AsString,
            FReader.FieldByName('OLDCITEMNAME').AsString,
            FReader.FieldByName('NEWCITEMCODE').AsString,
            FReader.FieldByName('NEWCITEMNAME').AsString,
            FReader.FieldByName('CHANGETYPE').AsString,
            FReader.FieldByName('PENALTYPE').AsString,
            FReader.FieldByName('PENAL').AsString,
            FReader.FieldByName('DEPOSITTYPE').AsString,
            FReader.FieldByName('FACIITEM').AsString] );
        FWriter.ExecSQL;
        FReader.Next
      end;
    FReader.Close;

    //CD078G
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078G WHERE BPCODE=''%s'' ',
        [ TableOwner, BPCode.Value ] );
    FReader.Open;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
         ' INSERT INTO %s.CD078G ( BPCODE,CitemCode,CitemName,           ' +
         '   ServiceType,FaciItem,SalePointcode,                         ' +
          '   SalePointName,InstCodeStr )                                ' +
         ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                         ' +
          '       ''%s'', ''%s'', ''%s'',''%s''   )                      ' ,
        [ TableOwner,
        FCloneCodeNo,
        FReader.FieldByName( 'CitemCode' ).AsString,
        FReader.FieldByName( 'CitemName' ).AsString,
        FReader.FieldByName( 'ServiceType' ).AsString,
        FReader.FieldByName( 'FaciItem' ).AsString,
        FReader.FieldByName( 'SalePointCode' ).AsString,
        FReader.FieldByName( 'SalePointName').AsString,
        FReader.FieldByName( 'InstCodeStr').AsString ] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078G1
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078G1 WHERE BPCODE=''%s'' ',
        [TableOwner, BPCode.Value] );
    FReader.Open;
    while not FReader.Eof do
    begin
      aCD078G1Step := GetCD078G1Step( TableOwner );
      FWriter.Close;
      FWriter.SQL.Text := Format (
          ' INSERT INTO %s.CD078G1 ( BPCODE,CitemCode,                    ' +
          '   ServiceType,FaciItem,SalePointcode,SalePointName,           ' +
          '   StepNo,Description,MasterSale,DiscountAmt )                 ' +
          ' VALUES ( ''%s'', ''%s'' ,''%s'',                              ' +
          '       ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
          '       ''%s'', ''%s'' )                                        ' ,
        [ TableOwner,
         FCloneCodeNo ,
         FReader.FieldByName( 'CitemCode' ).AsString ,
         FReader.FieldByName( 'ServiceType' ).AsString ,
         FReader.FieldByName( 'FaciItem' ).AsString ,
         FReader.FieldByName( 'SalePointCode' ).AsString ,
         FReader.FieldByName( 'SalePointName' ).AsString ,
         aCD078G1Step,
         FReader.FieldByName( 'Description' ).AsString ,
         FReader.FieldByName( 'MasterSale' ).AsString ,
         FReader.FieldByName( 'DiscountAmt' ).AsString ] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078I
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078I WHERE BPCODE=''%s'' ',
        [ TableOwner, BPCode.Value ] );
    FReader.Open;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
         ' INSERT INTO %s.CD078I ( BPCODE,CitemCode,                     ' +
         '   ServiceType,FaciItem,DiscountCitemCode,                     ' +
          '   CASID,DiscountAmt,InstCodeStr )                            ' +
         ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                         ' +
          '       ''%s'', ''%s'', ''%s'',''%s''   )                      ' ,
        [ TableOwner,
        FCloneCodeNo,
        FReader.FieldByName( 'CitemCode' ).AsString,
        FReader.FieldByName( 'ServiceType' ).AsString,
        FReader.FieldByName( 'FaciItem' ).AsString,
        FReader.FieldByName( 'DiscountCitemCode' ).AsString,
        FReader.FieldByName( 'CASID').AsString,
        FReader.FieldByName( 'DiscountAmt').AsString,
        FReader.FieldByName( 'InstCodeStr').AsString ] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078J
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078J WHERE BPCODE=''%s'' ',
        [ TableOwner, BPCode.Value ] );
    FReader.Open;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
         ' INSERT INTO %s.CD078J ( BPCODE,PRODUCTCODE,                   ' +
         '   PRODUCTNAME,SERVICETYPE,FACIITEM,                           ' +
          '   INSTCODESTR )                                              ' +
         ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                         ' +
          '       ''%s'', ''%s''               )                         ' ,
        [ TableOwner,
        FCloneCodeNo,
        FReader.FieldByName( 'PRODUCTCODE' ).AsString,
        FReader.FieldByName( 'PRODUCTNAME' ).AsString,
        FReader.FieldByName( 'SERVICETYPE' ).AsString,
        FReader.FieldByName( 'FACIITEM' ).AsString,
        FReader.FieldByName( 'INSTCODESTR').AsString ] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078K
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078K WHERE BPCODE=''%s'' ',
        [ TableOwner, BPCode.Value ] );
    FReader.Open;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
         ' INSERT INTO %s.CD078K ( BPCODE,CITEMCODE,                   ' +
         '   MASTERITEM,AMOUNT,IFRSPERCENT,                            ' +
          '   IFRSAMT )                                                ' +
         ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                       ' +
          '       ''%s'', ''%s''               )                       ' ,
        [ TableOwner,
        FCloneCodeNo,
        FReader.FieldByName( 'CITEMCODE' ).AsString,
        FReader.FieldByName( 'MASTERITEM' ).AsString,
        FReader.FieldByName( 'AMOUNT' ).AsString,
        FReader.FieldByName( 'IFRSPERCENT' ).AsString,
        FReader.FieldByName( 'IFRSAMT').AsString ] );
      FWriter.ExecSQL;
      FReader.Next
    end;

    //CD078L
    FReader.Close;
    FReader.SQL.Text := Format (
        ' SELECT * FROM %s.CD078L WHERE BPCODE=''%s'' ',
        [ TableOwner, BPCode.Value ] );
    FReader.Open;
    while not FReader.Eof do
    begin
      FWriter.Close;
      FWriter.SQL.Text := Format (
         ' INSERT INTO %s.CD078L ( BPCODE,MONTH,                           ' +
         '   MASTERSERVICETYPE,CASHADJ,SUBTOTAL,                           ' +
         '   CMAMT,CMTAX,CMADJ ,                                           ' +
         '   DTVAMT,DTVTAX,DTVADJ,                                         ' +
         '   PRVAMT,PRVTAX,PRVADJ)                                         ' +
         ' VALUES ( ''%s'', ''%s'' , ''%s'', ''%s'',                       ' +
         '       ''%s'', ''%s'',''%s'', ''%s'',                            ' +
         '       ''%s'', ''%s'',''%s'', ''%s'',                            ' +
         '       ''%s'', ''%s''  ) ',
        [ TableOwner,
        FCloneCodeNo,
        FReader.FieldByName( 'MONTH' ).AsString,
        FReader.FieldByName( 'MASTERSERVICETYPE' ).AsString,
        FReader.FieldByName( 'CASHADJ' ).AsString,
        FReader.FieldByName( 'SUBTOTAL' ).AsString,
        FReader.FieldByName( 'CMAMT').AsString ,
        FReader.FieldByName( 'CMTAX').AsString ,
        FReader.FieldByName( 'CMADJ').AsString ,
        FReader.FieldByName( 'DTVAMT').AsString ,
        FReader.FieldByName( 'DTVTAX').AsString ,
        FReader.FieldByName( 'DTVADJ').AsString ,
        FReader.FieldByName( 'PRVAMT').AsString ,
        FReader.FieldByName( 'PRVTAX').AsString ,
        FReader.FieldByName( 'PRVADJ').AsString,
        FReader.FieldByName( 'STOPFLAG').AsString  ]
         );
      FWriter.ExecSQL;
      FReader.Next
    end;

    FReader.Close;
    FWriter.Close;


    DBController.DataConnection.CommitTrans;
    InfoMsg( '???s?????C');
  except
    on E: Exception do
    begin
      DBController.DataConnection.RollbackTrans;
      ErrorMsg( Format( '?s??????, ???]:%s?C', [E.Message] ) );
    end;
  end;
  Screen.Cursor := crDefault;
end;

{ -------------------------------------------------------------------------------------- }

procedure TfrmCD078B8.btnSaveClick(Sender: TObject);
begin
  if DataOk then
  begin
    CopyNew( DBController.LoginInfo.DbAccount );
    ModalResult := mrOk;
  end;
end;

{ -------------------------------------------------------------------------------------- }

function TfrmCD078B8.DataOk: Boolean;
var
  aErrMsg: String;
  aControl: TWinControl;
begin
  aErrMsg := EmptyStr;
  aControl := nil;
  try
    if BPCode.Value = EmptyStr then
      begin
        aErrMsg := '?????J?i???X???~?N?X?j?C';
        aControl := BPCode.CodeNo;
        Exit;
      end;
    if radCopyNew.Checked then
    begin
      if edtNewCodeNo.Text = EmptyStr then
        begin
          aErrMsg := '?????J?i?s???X???~?N?X?j?C';
          aControl := edtNewCodeNo;
          Exit;
        end;
      if edtNewCodeName.Text = EmptyStr then
        begin
          aErrMsg := '?????J?i?s???X???~?W???j?C';
          aControl := edtNewCodeName;
          Exit;
        end;
      FCloneCodeNo := edtNewCodeNo.Text;
      FCloneDescription := edtNewCodeName.Text;
      FReader.Close;
      FReader.SQL.Text := Format (
          ' SELECT CODENO FROM %s.CD078 WHERE CODENO = ''%S'' ',
          [DBController.LoginInfo.DbAccount, FCloneCodeNo]);
      FReader.Open;
      if not FReader.Eof then
      begin
        aErrMsg := Format( '?s?N?X ''%s'' ?w?g?s?b?F?C',[FCloneCodeNo]);
        aControl := edtNewCodeNo;
        Exit;
      end;
    end;
    if radCopyExist.Checked then
    begin
      if CopyToBPCode.Value = EmptyStr then
      begin
        aErrMsg := '?????J?i???s???u?f???X?N?X?j?C';
        aControl := CopyToBPCode.CodeNo;
        Exit;
      end;
      if BPCode.Value = CopyToBPCode.Value then
      begin
        aErrMsg := '?i?u?f???X?N?X?j?P?i???s???u?f???X?N?X?j???i?H???P?C';
        aControl := CopyToBPCode.CodeNo;
        Exit;
      end;
      FCloneCodeNo := CopyToBPCode.Value;
      FCloneDescription := BPCode.valueName;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
    if not Result then
    begin
      WarningMsg( aErrMsg );
      if Assigned( aControl ) then
        if aControl.CanFocus then aControl.SetFocus;
    end;
  end;
end;

function TfrmCD078B8.GetCD078G1Step(const AOwner: String): String;
begin
  FWriter.Close;
  FWriter.SQL.Text := Format(
    ' SELECT %s.S_CD078G1.NEXTVAL FROM DUAL ', [AOwner] );
  FWriter.Open;
  Result := FWriter.Fields[0].AsString;
  FWriter.Close;
end;

end.
