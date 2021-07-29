unit frmSO8B10U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Grids, DBGrids, ComCtrls, Buttons, StdCtrls, DBCtrls,
  Mask, DB ,DBClient;

type
  TfrmSO8B10 = class(TForm)
    pnlEdit: TPanel;
    Splitter1: TSplitter;
    pnlButton: TPanel;
    btnExit: TSpeedButton;
    btnSave: TSpeedButton;
    btnCancel: TSpeedButton;
    btnDelete: TSpeedButton;
    btnEdit: TSpeedButton;
    btnInsert: TSpeedButton;
    Label2: TLabel;
    Label3: TLabel;
    dtxCharge: TDBText;
    btnQueryCharge: TBitBtn;
    Label1: TLabel;
    dedRefNo: TDBEdit;
    Label4: TLabel;
    dsrSO120: TDataSource;
    dsrCodeCD039: TDataSource;
    cobComp: TComboBox;
    dsrCodeCD019: TDataSource;
    dsrCodeCD032: TDataSource;
    cobPayUnit: TComboBox;
    lblStatus: TLabel;
    gbxFirst: TGroupBox;
    Label5: TLabel;
    Label6: TLabel;
    dedFirstCreditCard1: TDBEdit;
    dedFirstNotCreditCard1: TDBEdit;
    lblUnit1: TLabel;
    lblUnit2: TLabel;
    cobPromote: TComboBox;
    Label11: TLabel;
    dsrCodeCD042: TDataSource;
    Label7: TLabel;
    dedChannelViewDays: TDBEdit;
    GroupBox2: TGroupBox;
    Label9: TLabel;
    Label10: TLabel;
    lblUnit3: TLabel;
    lblUnit4: TLabel;
    dedFirstCreditCard2: TDBEdit;
    dedFirstNotCreditCard2: TDBEdit;
    GroupBox3: TGroupBox;
    Label13: TLabel;
    Label14: TLabel;
    lblUnit5: TLabel;
    lblUnit6: TLabel;
    dedOtherCreditCard2: TDBEdit;
    dedOtherNotCreditCard2: TDBEdit;
    pnlGrid: TPanel;
    StatusBar1: TStatusBar;
    dgrMain: TDBGrid;
    procedure FormShow(Sender: TObject);
    procedure CurrentCursorAndCdsCount;
    procedure btnExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnQueryChargeClick(Sender: TObject);
    procedure dsrSO120DataChange(Sender: TObject; Field: TField);
    procedure btnCancelClick(Sender: TObject);
    procedure btnInsertClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure cobPayUnitChange(Sender: TObject);

  private
    { Private declarations }
    procedure mainDataChange;
    procedure ActionAppend;
    function IsDataOk : Boolean;
    function checkPercentOrAmount(sI_Percent_Amount : String) : Boolean;
    procedure ChangeState;
    function checkOnlyPayCode(sI_PayCode1,sI_PayCode2,sI_PayCode3,sI_PayCode4,sI_PayCode5 : String) : Integer;
  public
    { Public declarations }
    sG_CompCode : String;
    procedure changeUnitCaption(nI_Unit:Integer);
  end;

var
  frmSO8B10: TfrmSO8B10;

implementation

uses dtmMain1U, UCommonU, frmDbSelectu, Ustru, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8B10.FormShow(Sender: TObject);
begin
    self.Caption := frmMainMenu.setCaption('SO8B10','[�����p�⤧������/���B�N�X��]');

    //�D����ܰʮ�,�e����Ƹ���ܰ�
    mainDataChange;
end;

procedure TfrmSO8B10.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8B10.FormCreate(Sender: TObject);
var
    sL_PromoteCode : String;
begin
    //��� CC&B (VB) �ǨӪ����q�O�Ѽ�
    sG_CompCode := frmMainMenu.sG_CompCode;
    {
    //�Y���`���q�h��� ini �ɩҦ������q�O,��ܩ� ComboBox
    if dtmMain1.sG_IsHeadQuarters = 'Y' then
    begin
        dtmMain1.getAllCompDataSet;
        dsrCodeCD039.DataSet := dtmMain1.cdsCodeCD039Client;
    end;
    }
    //���q�O
    TUCommonFun.AddObjectToComboBox(cobComp.Items, dsrCodeCD039.DataSet,NOT INSERT_NO_DATA_ITEM,'CodeNo','Description');
    TUCommonFun.setComboDefaultNdx(cobComp, sG_CompCode);

    //�~�ȧO
    sL_PromoteCode := dsrSO120.DataSet.FieldByName('PROMOTECODE').AsString;
    TUCommonFun.AddObjectToComboBox(cobPromote.Items, dsrCodeCD042.DataSet,INSERT_NO_DATA_ITEM,'CodeNo','Description');
    TUCommonFun.setComboDefaultNdx(cobPromote, sL_PromoteCode);


    //��ܵ{�����A
    ChangeState;

    dtmMain1.getSO120Data(sG_CompCode);
end;

procedure TfrmSO8B10.btnQueryChargeClick(Sender: TObject);
var
    sL_SelectedFieldName : String;
begin
    dtmMain1.getCodeCD019;

    if (dsrSO120.DataSet.State=dsBrowse) and (dsrSO120.DataSet.RecordCount =0) then
      dsrSO120.DataSet.Append;

    sL_SelectedFieldName := UpperCase(dgrMain.SelectedField.FieldName);

    if SelectRecord('���I�怜�O����', dsrCodeCD019.DataSet, 'CODENO;DESCRIPTION','') = mrOk then
    begin
      if not (dsrSO120.DataSet.State in [dsEdit, dsInsert]) then
        dsrSO120.DataSet.Edit;
      dsrSO120.DataSet.FieldByName('CODENO').AsString := dsrCodeCD019.DataSet.FieldByName('CodeNo').AsString;
      dsrSO120.DataSet.FieldByName('DESCRIPTION').AsString := dsrCodeCD019.DataSet.FieldByName('DESCRIPTION').AsString;
    end;
end;

procedure TfrmSO8B10.dsrSO120DataChange(Sender: TObject; Field: TField);
var
    sL_PayCode1,sL_PayCode2,sL_PayCode3,sL_PayCode4,sL_PayCode5 : String;
begin
    //�D����ܰʮ�,�e����Ƹ���ܰ�
    mainDataChange;
end;

procedure TfrmSO8B10.mainDataChange;
var
    nL_PayUnit : Integer;
    sL_PayUnit,sL_CompCode,sL_PromoteCode : String;
begin
    //�D����ܰʮ�,�e����Ƹ���ܰ�
    if dsrSO120.DataSet.Active then
    begin
      //���s�����A��
      if dsrSO120.DataSet.State = dsBrowse then
      begin
        sL_PayUnit := dsrSO120.DataSet.FieldByName('PAYUNIT').AsString;
        if sL_PayUnit = '1' then
          cobPayUnit.ItemIndex := 0
        else if sL_PayUnit = '2' then
          cobPayUnit.ItemIndex := 1;

        if sL_PayUnit <> '' then
          nL_PayUnit := StrToInt(sL_PayUnit);
        changeUnitCaption(nL_PayUnit);

        sL_CompCode := dsrSO120.DataSet.FieldByName('COMPCODE').AsString;
        TUCommonFun.setComboDefaultNdx(cobComp, sL_CompCode);

        sL_PromoteCode := dsrSO120.DataSet.FieldByName('PROMOTECODE').AsString;
        if sL_PromoteCode = '' then
          TUCommonFun.setComboDefaultNdx(cobPromote, '0')
        else
          TUCommonFun.setComboDefaultNdx(cobPromote, sL_PromoteCode);
      end;
    end;

    CurrentCursorAndCdsCount;
end;

procedure TfrmSO8B10.btnCancelClick(Sender: TObject);
begin
    with dsrSO120.DataSet do
    begin
        if State = dsInactive then Exit;
         Cancel;
         ChangeState;
         
    end;
end;

procedure TfrmSO8B10.ActionAppend;
var
    nL_NewTotalRecordCount : Integer;
begin
    nL_NewTotalRecordCount := dtmMain1.qrySO120.RecordCount + 1;


    with dsrSO120.DataSet do
    begin
        DisableControls;
        if State = dsInactive then
            Active := True;
        Append;

//        dtxCharge.Caption := '';
        cobPayUnit.ItemIndex := 0;
//        dedRefNo.Text := '';

        TUCommonFun.setComboDefaultNdx(cobPromote, '0');

        ChangeState;
        cobComp.Enabled := true;
        btnQueryCharge.Enabled := true;
        EnableControls;
    end;

    //�s�W�ɦC�ƥ[ 1 ���
    StatusBar1.Panels[0].Text:='                                                '
    + '                                                                          '+
    IntToStr(nL_NewTotalRecordCount) + ' ' + '/' + ' ' + IntToStr(nL_NewTotalRecordCount);

    //�������w�]���ʤ���
    cobPayUnit.ItemIndex := 0;

    //�w�]�� VB �ǨӪ����q�O
    TUCommonFun.setComboDefaultNdx(cobComp, sG_CompCode);

end;


procedure TfrmSO8B10.btnInsertClick(Sender: TObject);
begin
    pnlGrid.Enabled := FALSE;
    ActionAppend;
end;

procedure TfrmSO8B10.btnEditClick(Sender: TObject);
begin
    pnlGrid.Enabled := FALSE;
    mainDataChange;
    with dsrSO120.DataSet do
    begin
        if State = dsInactive then Exit;
            Edit;

        cobComp.Enabled := false;
        btnQueryCharge.Enabled := false;


        ChangeState;
    end;
    cobPayUnit.SetFocus;
end;

procedure TfrmSO8B10.btnDeleteClick(Sender: TObject);
begin
  with dsrSO120.DataSet do
  begin
    if State = dsInactive then Exit;
     if MessageDlg('�O�_�T�{�R��?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
       Exit;
     Delete;
     ChangeState;
  end;
end;

procedure TfrmSO8B10.btnSaveClick(Sender: TObject);
var
    sL_CompCode,sL_CodeNo,sL_CodeDesc,sL_PayUnit ,sL_CurDateTime : String;
    nL_PayCodeNo,nL_PayUnit,nL_PayUnitComboIndex : Integer;
    bL_IsOnlyData : Boolean;
    sL_PromoteCode,sL_PromoteName,sL_DiscountCode,sL_DiscountName : String;
begin
    //�����b�s��e���s�W�έק粒���̫�@�����ɸ���
    cobPayUnit.SetFocus;

    if IsDataOk then
    begin
        sL_CurDateTime := DateTimeToStr(Now);
        sL_CompCode := dtmMain1.getCurCobRecID(cobComp);
        sL_CodeNo := dsrSO120.DataSet.FieldByName('CODENO').AsString;
        sL_PromoteCode := dtmMain1.getCurCobRecID(cobPromote);
        //���o�~�Ȭ��ʥN�X�۹������W��,�u�f��k�N�X���u�f��k�W��
        dtmMain1.getCD024Data(sL_PromoteCode,sL_PromoteName,sL_DiscountCode,sL_DiscountName);
        //sL_CodeDesc := dsrSO120.DataSet.FieldByName('DESCRIPTION').AsString;

        nL_PayUnitComboIndex := cobPayUnit.ItemIndex;

        if dtmMain1.qrySO120.State in [dsInsert] then //�s�W���A�ˬd�O�_����
        begin
          //�ˬd��ƬO�_����
          bL_IsOnlyData := dtmMain1.checkIsSO120OnlyData(sL_CompCode,sL_CodeNo,sL_PromoteCode,sL_DiscountCode);
          if not bL_IsOnlyData then
          begin
            MessageDlg('��ƭ��Ƥ����x�s!',mtError, [mbOK],0);
            dtmMain1.qrySO120.Cancel;
            ChangeState;
            ActionAppend;
            exit;
          end;
        end;

        with dsrSO120.DataSet do
        begin
            FieldByName('COMPCODE').AsInteger := StrToInt(sL_CompCode);

            //�Y�~�Ȭ��ʥN�X���O�� "�L" �ΥN�X�� "0"
            if sL_PromoteCode <> '0' then
            begin
              FieldByName('PROMOTECODE').AsInteger := StrToInt(sL_PromoteCode);
              FieldByName('PROMOTENAME').AsString := sL_PromoteName;
              FieldByName('DISCOUNTCODE').AsInteger := StrToInt(sL_DiscountCode);
              FieldByName('DISCOUNTNAME').AsString := sL_DiscountName;
            end;

            //FieldByName('CODENO').AsInteger := StrToInt(sI_CodeNo);
            //FieldByName('DESCRIPTION').AsString := sI_CodeDesc;

            if nL_PayUnitComboIndex = 0 then
            begin
                FieldByName('PAYUNIT').AsString := '1';
            end
            else
            begin
                FieldByName('PAYUNIT').AsString := '2';
            end;


            FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
            FieldByName('UPDTIME').AsString := sL_CurDateTime;

            dsrSO120.DataSet.Post;
        end;

        ChangeState;
    end;
end;

function TfrmSO8B10.IsDataOk: Boolean;
var
    sL_CodeNo,sL_ChannelViewDays : String;
    nL_Length,nL_PayUnitComboIndex : Integer;
    bL_CheckOk : Boolean;
    nL_PercentOrAmt1,nL_PercentOrAmt2,nL_PercentOrAmt3,nL_PercentOrAmt4 : Integer;
    sL_PercentOrAmt1,sL_PercentOrAmt2,sL_PercentOrAmt3,sL_PercentOrAmt4 : String;
begin

    sL_CodeNo := dsrSO120.DataSet.fieldByName('CodeNo').AsString;
    if sL_CodeNo = '' then
    begin
      if MessageDlg('�O�_���ۨӹq?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      begin
        MessageDlg('���O���ئW�٤��ର�ŭ�',mtError, [mbOK],0);
        btnQueryCharge.SetFocus;
        Result := false;
        exit;
      end
      else
      begin
       dsrSO120.DataSet.fieldByName('REF_NO').AsInteger := 9;
      end;
    end;

    //�W�D�ݦ������h�֤Ѥ~������,���ର�ŭ�
    sL_ChannelViewDays := dsrSO120.DataSet.fieldByName('ChannelViewDays').AsString;
    if sL_ChannelViewDays = '' then
    begin
      MessageDlg('�W�D�ݦ������h�֤Ѥ~���������ର�ŭ�',mtError, [mbOK],0);
      dedChannelViewDays.SetFocus;
      Result := false;
      exit;
    end;


    nL_PayUnitComboIndex := cobPayUnit.ItemIndex;
    //����
    if nL_PayUnitComboIndex = 0 then   //�Y���ʤ���h���p��100
    begin
      //�����H�Υd�ʤ��� -���s�H
      nL_PercentOrAmt1 := dsrSO120.DataSet.fieldByName('FIRSTCREDITCARD1').asInteger;
      if (nL_PercentOrAmt1 > 100) or (nL_PercentOrAmt1 < 0) then
      begin
        MessageDlg('������쬰�ʤ���,�G��>0, <100',mtError, [mbOK],0);
        dedFirstCreditCard1.SetFocus;
        Result := false;
        exit;
      end;

      //�����D�H�Υd�ʤ��� -���s�H
      nL_PercentOrAmt2 := dsrSO120.DataSet.fieldByName('FIRSTNOTCREDITCARD1').asInteger;
      if (nL_PercentOrAmt2 > 100) or (nL_PercentOrAmt2 < 0) then
      begin
        MessageDlg('������쬰�ʤ���,�G��>0, <100',mtError, [mbOK],0);
        dedFirstNotCreditCard1.SetFocus;
        Result := false;
        exit;
      end;


      //�����H�Υd�ʤ��� -���O��
      nL_PercentOrAmt1 := dsrSO120.DataSet.fieldByName('FIRSTCREDITCARD2').asInteger;
      if (nL_PercentOrAmt1 > 100) or (nL_PercentOrAmt1 < 0) then
      begin
        MessageDlg('������쬰�ʤ���,�G��>0, <100',mtError, [mbOK],0);
        dedFirstCreditCard2.SetFocus;
        Result := false;
        exit;
      end;

      //�����D�H�Υd�ʤ��� -���O��
      nL_PercentOrAmt2 := dsrSO120.DataSet.fieldByName('FIRSTNOTCREDITCARD2').asInteger;
      if (nL_PercentOrAmt2 > 100) or (nL_PercentOrAmt2 < 0) then
      begin
        MessageDlg('������쬰�ʤ���,�G��>0, <100',mtError, [mbOK],0);
        dedFirstNotCreditCard2.SetFocus;
        Result := false;
        exit;
      end;

      //�򦬫H�Υd�ʤ��� -���O��
      nL_PercentOrAmt3 := dsrSO120.DataSet.fieldByName('OTHERCREDITCARD2').asInteger;
      if (nL_PercentOrAmt3 > 100) or (nL_PercentOrAmt3 < 0) then
      begin
        MessageDlg('������쬰�ʤ���,�G��>0, <100',mtError, [mbOK],0);
        dedOtherCreditCard2.SetFocus;
        Result := false;
        exit;
      end;


      //�򦬫D�H�Υd�ʤ��� -���O��
      nL_PercentOrAmt4 := dsrSO120.DataSet.fieldByName('OTHERNOTCREDITCARD2').asInteger;
      if (nL_PercentOrAmt4 > 100) or (nL_PercentOrAmt4 < 0) then
      begin
        MessageDlg('������쬰�ʤ���,�G��>0, <100',mtError, [mbOK],0);
        dedOtherNotCreditCard2.SetFocus;
        Result := false;
        exit;
      end;

    end;

    Result := true;

end;



function TfrmSO8B10.checkPercentOrAmount(sI_Percent_Amount: String): Boolean;
var
    nL_PayUnit,nL_TotalLength,nL_PositiveLength,nL_DecimalLength : Integer;
    fI_Percent_Amount : Double;
    L_StrList : TStringList;
    sL_Positive,sL_Decimal : String;
begin
    //�̦���������M�w����J�ʤ���Ϊ��B
    nL_PayUnit := cobPayUnit.ItemIndex;

    nL_TotalLength := Length(sI_Percent_Amount);
    if (nL_TotalLength > 11) Or (StrToFloat(sI_Percent_Amount) < 0)then
    begin
        if nL_PayUnit = 0 then       //������쬰�ʤ���
            MessageDlg('�ʤ��񤣯�W�L���� 2 ���,�p�� 2 ���',mtError, [mbOK],0)
        else if nL_PayUnit = 1 then  //������쬰��
            MessageDlg('���B����W�L���� 8 ���,�p�� 2 ���',mtError, [mbOK],0);

        Result := false;
        exit;
    end;


    L_StrList := TUStr.ParseStrings(sI_Percent_Amount,'.');
    sL_Positive := L_StrList.Strings[0];
    sL_Decimal := L_StrList.Strings[1];
    nL_PositiveLength := Length(sL_Positive);
    nL_DecimalLength := Length(sL_Decimal);


    if nL_PayUnit = 0 then       //������쬰�ʤ���
    begin
        if (nL_PositiveLength > 2) OR (nL_DecimalLength > 2)  then
        begin
            MessageDlg('�ʤ��񤣯�W�L���� 2 ���,�p�� 2 ���',mtError, [mbOK],0);
            Result := false;
            exit;
        end;
    end
    else if nL_PayUnit = 1 then  //������쬰��
    begin
        if (nL_PositiveLength > 8) OR (nL_DecimalLength > 2)  then
        begin
            MessageDlg('���B����W�L���� 2 ���,�p�� 2 ���',mtError, [mbOK],0);
            Result := false;
            exit;
        end;
    end;
end;

procedure TfrmSO8B10.cobPayUnitChange(Sender: TObject);
var
    nL_PayUnit : Integer;
    sL_Msg : String;
begin
    nL_PayUnit := cobPayUnit.ItemIndex;

    //if nL_PayUnit = 0 then       //������쬰�ʤ���
    //    sL_Msg := '�����������N�|�ܴ��Ҧ��I�ڤ覡';

    {
    if MessageDlg('�����������N�|�ܴ��Ҧ��I�ڤ覡?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        Exit;
    }
    changeUnitCaption(nL_PayUnit+1);

end;


procedure TfrmSO8B10.ChangeState;
begin

    with dsrSO120.DataSet do
    begin
      if state in [dsInactive] then
        Exit
      else if state in [dsEdit, dsInsert] then
      begin
        btnCancel.Enabled := TRUE;
        btnSave.Enabled := TRUE;
        btnDelete.Enabled := FALSE;
        btnEdit.Enabled := FALSE;
        btnExit.Enabled := FALSE;
        btnInsert.Enabled := FALSE;
        pnlEdit.Enabled := TRUE;
        //pnlGrid.Enabled := FALSE;
      end
      else
      begin
        if RecordCount>0 then
        begin
          btnEdit.Enabled := TRUE;
          btnDelete.Enabled := TRUE;
        end
        else
        begin
          btnEdit.Enabled := FALSE;
          btnDelete.Enabled := FALSE;
        end;
        btnCancel.Enabled := FALSE;
        btnSave.Enabled := FALSE;
        btnInsert.Enabled := TRUE;
        btnExit.Enabled := TRUE;
        pnlEdit.Enabled := FALSE;
        pnlGrid.Enabled := TRUE;
       end;
     end;
     //down, �]�w�ާ@�Ҧ������


     with dsrSO120.DataSet do
     begin
       if state in [dsInsert] then
         lblStatus.Caption := '�s�W'
       else if state in [dsEdit] then
         lblStatus.Caption := '�ק�'
       else
         lblStatus.Caption := '�s��';
     end;
     //up, �]�w�ާ@�Ҧ������

end;




function TfrmSO8B10.checkOnlyPayCode(sI_PayCode1,sI_PayCode2,sI_PayCode3,sI_PayCode4,sI_PayCode5 : String): Integer;
var
    L_PayCodeList,L_PayCodeListTemp : TStringList;
    ii,nL_Ndx : Integer;
begin
    L_PayCodeList := TStringList.Create;
    L_PayCodeListTemp := TStringList.Create;
    L_PayCodeList.Add(sI_PayCode1);
    L_PayCodeList.Add(sI_PayCode2);
    L_PayCodeList.Add(sI_PayCode3);
    L_PayCodeList.Add(sI_PayCode4);
    L_PayCodeList.Add(sI_PayCode5);

    for ii:=0 to 4 do
    begin
        if L_PayCodeList.Strings[ii] <> '0' then   //���p�I�ڤ覡���O��ܵL
        begin
            if ii = 0 then
            begin
                L_PayCodeListTemp.Add(L_PayCodeList.Strings[ii]);
            end
            else
            begin
                nL_Ndx := L_PayCodeListTemp.IndexOf(L_PayCodeList.Strings[ii]);
                if nL_Ndx = -1 then
                begin
                    L_PayCodeListTemp.Add(L_PayCodeList.Strings[ii]);
                    Result := -1;
                end
                else
                begin
                    Result := ii + 1;   //�^�ǲĴX�ӥI�ڤ覡����
                    exit;
                end;
            end;
        end;
    end;
end;

procedure TfrmSO8B10.CurrentCursorAndCdsCount;
begin
    StatusBar1.Panels[0].Text:='                                                '
    + '                                                                          '+
    IntToStr(dtmMain1.qrySO120.RecNo) + ' ' + '/' + ' ' + IntToStr(dtmMain1.qrySO120.RecordCount);

end;


procedure TfrmSO8B10.changeUnitCaption(nI_Unit: Integer);
begin
    case nI_Unit of
      1:        //������쬰�ʤ���
        begin
            lblUnit1.Caption := '%';
            lblUnit2.Caption := '%';
            lblUnit3.Caption := '%';
            lblUnit4.Caption := '%';
            lblUnit5.Caption := '%';
            lblUnit6.Caption := '%';
        end;
      2:  //������쬰���B
        begin
            lblUnit1.Caption := '��';
            lblUnit2.Caption := '��';
            lblUnit3.Caption := '��';
            lblUnit4.Caption := '��';
            lblUnit5.Caption := '��';
            lblUnit6.Caption := '��';
        end;
    end;

end;

end.
