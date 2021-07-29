unit frmSnapShotU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, DBGrids, DB, Buttons, fraChineseYMDU ,DBClient;

type
  TfrmSnapShot = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    edtAlias: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    edtUserName: TEdit;
    edtPassword: TEdit;
    fraChineseYMD1: TfraChineseYMD;
    Label4: TLabel;
    btnOK: TBitBtn;
    btnClose: TButton;
    dsrSO030A: TDataSource;
    dsrSO509B: TDataSource;
    Panel3: TPanel;
    DBGrid1: TDBGrid;
    lblCompName: TLabel;
    Panel5: TPanel;
    DBGrid2: TDBGrid;
    btnReCompute: TButton;
    btnAllReCompute: TBitBtn;
    procedure btnCloseClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure btnReComputeClick(Sender: TObject);
    procedure btnAllReComputeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    nG_TotalCustCounts : Integer;
    sG_EngDate : String;
    function isDataOk : Boolean;
  public
    { Public declarations }
    nG_Counts : Integer;
    dG_CurDateTime : TDateTime;
    procedure filterCdsSO001(sI_Filter : String);
  end;

var
  frmSnapShot: TfrmSnapShot;

implementation

uses Ustru, dtmMainU, UCommonU;

{$R *.dfm}

procedure TfrmSnapShot.btnCloseClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSnapShot.btnOKClick(Sender: TObject);
var
    sL_Alias,sL_UserName,sL_Password,sL_Date : String;
    sL_CompCode,sL_CompName,sL_Filter,sL_DbOwner : String;
begin
    if isDataOk then
    begin
      //���浥�ݪ��A
      TUCommonFun.setWaitingCursor;

      sL_Alias :=  Trim(edtAlias.Text);

      //�H�����s���N��
      sL_UserName := DEFAULT_DB_OWNER_NAME;
      sL_Password := DEFAULT_DB_OWNER_PASSWORD;

      sL_Date := fraChineseYMD1.getYMD;
//showmessage('1' + '=' + sL_Alias + '=' + sL_UserName + '=' + sL_Password);
      dtmMain.connectToDB(sL_Alias,sL_UserName,sL_Password);
//showmessage('2');

      if TClientDataSet(dtmMain.cdsSO030A).IndexName = 'TmpIndex' then
        TClientDataSet(dtmMain.cdsSO030A).DeleteIndex('TmpIndex');

      //�����{�b�ɶ�
      dG_CurDateTime := now;


      sG_EngDate := IntToStr((StrToInt(Copy(sL_Date,1,3))+1911)) + Copy(sL_Date,4,6);
//showmessage('3');
      //Select �Ҧ����q SnapShop ���������
      if dtmMain.getSnapShotLogData(sG_EngDate) then
      begin
        //Select ������ Table �Q Move SnapShot
        dtmMain.getMoveToSnapShotTable;

        sL_CompCode := dsrSO030A.DataSet.FieldByName('CompCode').AsString;
        sL_CompName := dsrSO030A.DataSet.FieldByName('DESCRIPTION').AsString;
        sL_DbOwner := dtmMain.getDbOwner(sL_CompCode);

        lblCompName.Caption := sL_CompName;
        btnReCompute.Enabled := true;
        btnAllReCompute.Enabled := true;

        //�p�� SO �P Snapshot ������
        dtmMain.computeSoAndSnapshotTableCounts(sL_DbOwner,sL_CompCode,sL_CompName,sG_EngDate);
      end
      else
      begin
//showmessage('5');
         lblCompName.Caption := '���q�W��';
         btnReCompute.Enabled := false;
         btnAllReCompute.Enabled := false;
         dtmMain.cdsSO509B.Close;
         dtmMain.cdsTempSO509B.EmptyDataSet;
         MessageDlg(QUERY_NO_RECORD,mtInformation, [mbOK],0);
      end;

      //�^�_�쪬�A
      TUCommonFun.setDefaultCursor;
    end;
end;

function TfrmSnapShot.isDataOk: Boolean;
begin
    if Trim(edtAlias.Text) = '' then
    begin
      MessageDlg('�п�JAlias',mtError, [mbOK],0);
      edtAlias.SetFocus;
      Result := false;
      exit;
    end;

{    
    if Trim(edtUserName.Text) = '' then
    begin
      MessageDlg('�п�JUserName',mtError, [mbOK],0);
      edtUserName.SetFocus;
      Result := false;
      exit;
    end;

    if Trim(edtPassword.Text) = '' then
    begin
      MessageDlg('�п�JPassword',mtError, [mbOK],0);
      edtPassword.SetFocus;
      Result := false;
      exit;
    end;
}

    if Trim(TUstr.replaceStr(fraChineseYMD1.getYMD,'/','')) = '' then
    begin
      MessageDlg('�п�J�d�ߤ��',mtError, [mbOK],0);
      fraChineseYMD1.mseYMD.SetFocus;
      Result := false;
      exit;
    end;
end;

procedure TfrmSnapShot.DBGrid1TitleClick(Column: TColumn);
var
    sL_ColumnName : String;
begin
    sL_ColumnName := Column.FieldName;

    if TClientDataSet(dtmMain.cdsSO030A).IndexName = 'TmpIndex' then
      TClientDataSet(dtmMain.cdsSO030A).DeleteIndex('TmpIndex');

    TClientDataSet(dtmMain.cdsSO030A).AddIndex('TmpIndex', sL_ColumnName, [ixCaseInsensitive],'','',0);
    TClientDataSet(dtmMain.cdsSO030A).IndexName := 'TmpIndex';
end;

procedure TfrmSnapShot.DBGrid1CellClick(Column: TColumn);
var
    sL_CompCode,sL_CompName,sL_Filter,sL_DbOwner : String;
begin
    //���浥�ݪ��A
    TUCommonFun.setWaitingCursor;

    sL_CompCode := dsrSO030A.DataSet.FieldByName('CompCode').AsString;
    sL_CompName := dsrSO030A.DataSet.FieldByName('DESCRIPTION').AsString;

    sL_DbOwner := dtmMain.getDbOwner(sL_CompCode);

    //�p�� SO �P Snapshot ������
    dtmMain.computeSoAndSnapshotTableCounts(sL_DbOwner,sL_CompCode,sL_CompName,sG_EngDate);

    lblCompName.Caption := sL_CompName;

    //�^�_�쪬�A
    TUCommonFun.setDefaultCursor;
end;

procedure TfrmSnapShot.filterCdsSO001(sI_Filter: String);
var
    nL_CustCounts : Integer;
begin

    with dtmMain.cdsSO001 do
    begin
      Filtered := false;
      Filter := sI_Filter;
      Filtered := true;

      nG_TotalCustCounts := 0;
      First;
      while not Eof do
      begin
        nL_CustCounts := FieldByName('Counts').AsInteger;
        nG_TotalCustCounts := nG_TotalCustCounts + nL_CustCounts;
        Next;
      end;

      First;
    end;


end;

procedure TfrmSnapShot.btnReComputeClick(Sender: TObject);
var
    sL_CompCode,sL_CompName,sL_DbOwner : String;
begin
    //���浥�ݪ��A
    TUCommonFun.setWaitingCursor;

    //�����{�b�ɶ�
    dG_CurDateTime := now;

    sL_CompCode := dsrSO030A.DataSet.FieldByName('CompCode').AsString;
    sL_CompName := dsrSO030A.DataSet.FieldByName('DESCRIPTION').AsString;
    sL_DbOwner := dtmMain.getDbOwner(sL_CompCode);

    //�R���Ӥ��q�Ӥ骺SO509B���
    dtmMain.deleteSO509B(sL_DbOwner,sL_CompCode,sG_EngDate);

    //���s�p��
    dtmMain.computeSO509B(sL_DbOwner,sL_CompCode,sL_CompName,sG_EngDate);

    //�^�_�쪬�A
    TUCommonFun.setDefaultCursor;
end;

procedure TfrmSnapShot.btnAllReComputeClick(Sender: TObject);
var
    ii : Integer;
    sL_CompCode,sL_CompName,sL_DbOwner : String;
begin
    //���浥�ݪ��A
    TUCommonFun.setWaitingCursor;
    
    //�����{�b�ɶ�
    dG_CurDateTime := now;

    dtmMain.getAllCompCode;

    for ii:=0 to dtmMain.G_CompCodeStrList.Count-1 do
    begin
      sL_CompCode := dtmMain.G_CompCodeStrList.Strings[ii];
      sL_CompName := dtmMain.getCompName(sL_CompCode);
      sL_DbOwner := dtmMain.getDbOwner(sL_CompCode);

      //�R���Ӥ��q�Ӥ骺SO509B���
      dtmMain.deleteSO509B(sL_DbOwner,sL_CompCode,sG_EngDate);

      //���s�p��
      dtmMain.computeSO509B(sL_DbOwner,sL_CompCode,sL_CompName,sG_EngDate);

      //showmessage(dtmMain.G_CompCodeStrList.Strings[ii]);
    end;

    //�^�_�쪬�A
    TUCommonFun.setDefaultCursor;
end;

procedure TfrmSnapShot.FormShow(Sender: TObject);
begin
    btnReCompute.Enabled := false;
    btnAllReCompute.Enabled := false;
end;

end.
