unit frmQryCD042U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, Buttons, ExtCtrls, DB, ADODB, ActnList;

type
  TfrmQryCD042 = class(TForm)
    Panel1: TPanel;
    EDT_QryCD042Str: TEdit;
    X_QryCD042: TBitBtn;
    btnQryCD042: TButton;
    Label4: TLabel;
    EDT_StartDate: TMaskEdit;
    Label5: TLabel;
    EDT_StopDate: TMaskEdit;
    chkStopFlag: TCheckBox;
    btnOK: TButton;
    btnCancel: TButton;
    ADOConnection1: TADOConnection;
    ActionList1: TActionList;
    Action1: TAction;
    procedure btnOKClick(Sender: TObject);
    procedure btnQryCD042Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure X_QryCD042Click(Sender: TObject);
    procedure Action1Execute(Sender: TObject);
    procedure EDT_StartDateExit(Sender: TObject);
    procedure EDT_StopDateExit(Sender: TObject);
    procedure EDT_StopDateEnter(Sender: TObject);
  private
    { Private declarations }
    FSQLWhere : string;
    FDBPassWord : String;
    FDBUserID : String;
    FDBName : String;
    FSelectCode : string;
    FSelectName : String;
  public
    { Public declarations }
    property sWhere : String read FSQLWhere;
  end;

var
  frmQryCD042: TfrmQryCD042;

implementation
uses
   frmMultiSelectU2,cbUtilis;
{$R *.dfm}

procedure TfrmQryCD042.btnOKClick(Sender: TObject);
  var aStartDate,aStopDate:string;
begin
  if Length(FSelectCode)>0 then
    FSQLWhere := ' And CodeNo In(' + FSelectCode + ')'
  else begin
    if not chkStopFlag.Checked then FSQLWhere :=' And StopFlag<>1';
    if EDT_StartDate.Text<>EmptyStr then
    begin
      aStartDate := TrimChar( EDT_StartDate.Text, ['/'] );
      FSQLWhere := ( FSQLWhere + Format( ' AND ActStartDate >= TO_DATE( ''%s'', ''YYYYMMDD'' ) ', [aStartDate] ) );
    end;
    if EDT_StopDate.Text<>EmptyStr then
    begin
      aStopDate := TrimChar( EDT_StopDate.Text , ['/'] );
      FSQLWhere := ( FSQLWhere + Format( ' AND ActStopDate <= TO_DATE( ''%s'', ''YYYYMMDD'' ) ', [aStopDate] ) );
    end;
  end;

  Self.ModalResult := mrOk;

end;

procedure TfrmQryCD042.btnQryCD042Click(Sender: TObject);
  var aWhere:string;
      aStartDate,aStopDate:string;
begin
  if not chkStopFlag.Checked then aWhere :=' And StopFlag<>1';
  if EDT_StartDate.Text<>EmptyStr then
  begin
    aStartDate := TrimChar( EDT_StartDate.Text, ['/'] );
    aWhere := ( aWhere + Format( ' AND ActStartDate >= TO_DATE( ''%s'', ''YYYYMMDD'' ) ', [aStartDate] ) );
  end;
  if EDT_StopDate.Text<>EmptyStr then
  begin
    aStopDate := TrimChar( EDT_StopDate.Text , ['/'] );
    aWhere := ( aWhere + Format( ' AND ActStopDate <= TO_DATE( ''%s'', ''YYYYMMDD'' ) ', [aStopDate] ) );
  end;
  frmMultiSelect2 :=TfrmMultiSelect2.Create(Application);
  try
    frmMultiSelect2.Connection :=ADOConnection1;
    frmMultiSelect2.KeyFields:='CODENO';
    frmMultiSelect2.KeyValues:= FSelectCode;
    frmMultiSelect2.DisplayFields:='CODENO,促銷活動代碼,Description,促銷活動名稱';
    frmMultiSelect2.ResultFields:='DESCRIPTION';

    frmMultiSelect2.SQL.Text := Format(
                            ' SELECT * FROM %S.CD042 ' +
                            '  WHERE ( DISCOUNTCODE IS NULL ) AND ( PRODUCTTYPE IS NOT NULL ) ' + aWhere +
                            '  ORDER BY CODENO', [FDBUserID] );


    if frmMultiSelect2.ShowModal=mrOk then
    begin
      FSelectCode := frmMultiSelect2.SelectedValue;
      FSelectName := frmMultiSelect2.SelectedDisplay;
      EDT_QryCD042Str.Text :=  FSelectName;
    end;
  finally
    frmMultiSelect2.Free;
  end;

end;

procedure TfrmQryCD042.FormCreate(Sender: TObject);
begin
  FDBPassWord := ParamStr( 7 );
  FDBUserID := ParamStr( 6 );
  FDBName := ParamStr( 8 );
  ADOConnection1.ConnectionString := 'Provider=MSDAORA.1;Password='+ FDBPassWord +';User ID='+FDBUserID+';Data Source='+FDBName;
end;

procedure TfrmQryCD042.btnCancelClick(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
end;

procedure TfrmQryCD042.X_QryCD042Click(Sender: TObject);
begin
  FSQLWhere:=EmptyStr;
  EDT_QryCD042Str.Clear;
end;

procedure TfrmQryCD042.Action1Execute(Sender: TObject);
begin
  btnOKClick(Sender);
end;

procedure TfrmQryCD042.EDT_StartDateExit(Sender: TObject);
  var aDate, aErrMsg: String;


begin
  aDate := TMaskEdit( Sender ).EditText;
  if not DateTextIsValidEx( aDate ) then
  begin
    aErrMsg := '您輸入的日期格式有誤。';
  end;
  if ( aErrMsg <> EmptyStr ) then
  begin
    WarningMsg( aErrMsg );
    if ( TMaskEdit( Sender ).CanFocus ) then
      TMaskEdit( Sender ).SetFocus;
  end else
  begin
    TMaskEdit( Sender ).EditText := aDate;
  end;

  
end;

procedure TfrmQryCD042.EDT_StopDateExit(Sender: TObject);
  var aDate, aErrMsg: String;
begin
  aDate := TMaskEdit( Sender ).EditText;
  if not DateTextIsValidEx( aDate ) then
  begin
    aErrMsg := '您輸入的日期格式有誤。';
  end;
  if ( aErrMsg <> EmptyStr ) then
  begin
    WarningMsg( aErrMsg );
    if ( TMaskEdit( Sender ).CanFocus ) then
      TMaskEdit( Sender ).SetFocus;
  end else
  begin
    TMaskEdit( Sender ).EditText := aDate;
  end;
end;

procedure TfrmQryCD042.EDT_StopDateEnter(Sender: TObject);
  var
  aDateTime1,aDateTime2 : TDateTime;
  aYear,aMonth,aDay : Word;
  aDayTable: PDayTable;
  aLastDay: Word;
begin
  try
    if EDT_StartDate.Text<>EmptyStr then
      TryStrToDate(EDT_StartDate.EditText,aDateTime1 );
    if EDT_StartDate.Text<>EmptyStr then
      TryStrToDate(EDT_StopDate.EditText,aDateTime2);
    if ((EDT_StartDate.Text<>EmptyStr) And (EDT_StopDate.Text=EmptyStr ))   then
    begin
      DecodeDate(aDateTime1,aYear,aMonth,aDay);
      aDayTable := @MonthDays[IsLeapYear(aYear)];
      aLastDay :=  aDayTable[aMonth];
      EDT_StopDate.EditText := Format( '%.4d/%.2d/%.2d', [aYear, aMonth, aLastDay] );
    end;
  finally
    EDT_StopDate.SelectAll;
  end;
end;

end.
