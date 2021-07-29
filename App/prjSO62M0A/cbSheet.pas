unit cbSheet;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Forms, ComObj, EXCEL_TLB, AdoDb;

var
  ExcelApp: OleVariant;
  procedure MakeCell(X, Y: Integer; ShowText: string; ColorIndex, AlignWhere: Integer);
  procedure MakeGutter(RowPos: Integer);
  procedure MergeCell(X1, Y1, X2, Y2: Integer);

procedure ExportToExcel(BPCode: string);

implementation

uses
  cbDataController;

procedure ExportToExcel(BPCode: string);
var
  aIndex, aIndex2, aIndex3, aColIndex, aRowIndex, aAmtPrice, aService, aMonthCount: Integer;
  aTotal, aMaxRowIndex, aInitRowIndex: Integer;
  aDescription: string;
  aList: TStringList;
  aDataReader: TAdoQuery;
begin
  aList := TStringList.Create;
  aDataReader := DataController.CodeReader;

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Workbooks.Add;
  ExcelApp.Visible := True;

  //����X BPCODE �����ǪA�����O
  aDataReader.Close;
  aDataReader.Sql.Text := Format ( 'SELECT DISTINCT SERVICETYPE  ' +
                                   'FROM %s.CD078A               ' +
                                   'WHERE BPCODE=''%s''          ' ,
                                   [DataController.LoginInfo.DbAccount,
                                    BPCODE] );
  aDataReader.Open;
  while not aDataReader.Eof do
    begin
      aList.Add(aDataReader.FieldByName('SERVICETYPE').AsString);
      aDataReader.Next;
    end;

  //�Ĥ@��
  MakeCell(1,1,'�A�����O',34,xlCenter);
  aColIndex := 2;
  for aIndex := 0 to aList.Count-1 do
    begin
      MergeCell(aColIndex,1,aColIndex+4,1);
      MakeCell(aColIndex,1,aList[aIndex],34,xlCenter);
      //�ĤG��A�ΨӰO���C�� ServideType ���X�Ӷ��ءA�������ץX�� EXCEL �A�M��
//      MergeCell(aColIndex,2,aColIndex+4,2);
//      MakeCell(aColIndex,2,'0',-1,xlCenter);
      aColIndex := aColIndex + 5;
    end;
  MakeCell(aColIndex,1,'�P��'+#10+'�`�p',34,xlCenter);
  MakeCell(aColIndex+1,1,'�u�f��'+#10+'�`�p',34,xlCenter);
  MakeCell(aColIndex+2,1,'���B�t��'+#10+'�`�p',34,xlCenter);

  MakeGutter(2);

  //�զX���~���Y�C
  MakeCell(1,3,'�զX���~',34,xlCenter);
  aColIndex := 2;
  for aIndex := 0 to aList.Count-1 do
    begin
      MakeCell(aColIndex,3,'���O����',34,xlLeft);
      MakeCell(aColIndex+1,3,'�j��'+#10+'���',34,xlCenter);
      MakeCell(aColIndex+2,3,'�P��',34,xlCenter);
      MakeCell(aColIndex+3,3,'�u�f��',34,xlCenter);
      MakeCell(aColIndex+4,3,'���B'+#10+'�t��',34,xlCenter);
      aColIndex := aColIndex + 5;
    end;
  MakeCell(aColIndex,3,'�P��'+#10+'�`�p',34,xlCenter);
  MakeCell(aColIndex+1,3,'�u�f��'+#10+'�`�p',34,xlCenter);
  MakeCell(aColIndex+2,3,'���B�t��'+#10+'�`�p',34,xlCenter);

  //�զX���~����
  aDataReader.Close;
  aDataReader.Sql.Text := Format ( 'SELECT DESCRIPTION  ' +
                                   'FROM %s.CD078       ' +
                                   'WHERE CODENO=''%s'' ' ,
                                   [DataController.LoginInfo.DbAccount,
                                    BPCODE] );
  aDataReader.Open;
  aDescription := aDataReader.FieldByName('DESCRIPTION').AsString;
  aColIndex := 2;
  aMaxRowIndex := -1;
  for aIndex := 0 to aList.Count-1 do
    begin
      aRowIndex := 4;
      aDataReader.Close;
      aDataReader.Sql.Text := Format ( 'SELECT CITEMNAME, BUNDLEMON, MONTHAMT,      ' +
                                       'Mon1, Mon2, Mon3, Mon4, Mon5, Mon6,         ' +
                                       'Mon7, Mon8, Mon9, Mon10, Mon11, Mon12,      ' +
                                       'DiscountAmt1, DiscountAmt2, DiscountAmt3,   ' +
                                       'DiscountAmt4, DiscountAmt5, DiscountAmt6,   ' +
                                       'DiscountAmt7, DiscountAmt8, DiscountAmt9,   ' +
                                       'DiscountAmt10, DiscountAmt11, DiscountAmt12 ' +
                                       'FROM %s.CD078A                              ' +
                                       'WHERE BPCODE=''%s''                         ' +
                                       'AND SERVICETYPE=''%s''                      ' ,
                                       [DataController.LoginInfo.DbAccount,
                                        BPCODE, aList[aIndex]] );
      aDataReader.Open;
      while not aDataReader.Eof do
        begin
          MakeCell(aColIndex,aRowIndex,aDataReader.FieldByName('CITEMNAME').AsString,-1,xlLeft);
          MakeCell(aColIndex+1,aRowIndex,aDataReader.FieldByName('BUNDLEMON').AsString,-1,xlRight);

          //���q���u�f��Ƥ��u�f���B�X�p
          aMonthCount := 0;
          aService := 0;
          for aIndex2 :=1 to 12 do
            begin
              aMonthCount := aMonthCount +
                             aDataReader.FieldByName('Mon'+IntToStr(aIndex2)).AsInteger;
              aService := aService +
                          aDataReader.FieldByName('DiscountAmt'+IntToStr(aIndex2)).AsInteger;
            end;
          if aDataReader.FieldByName('BUNDLEMON').AsInteger > 0 then
            begin
              //�j����� <> 0
              //�Y�u�f��Ƥp��j����ơA�n (�j�����-�u�f���) * �C��P��
              //�~�O�u�����u�f���B
              aService := aService + (aDataReader.FieldByName('BUNDLEMON').AsInteger - aMonthCount) *
                          aDataReader.FieldByName('MONTHAMT').AsInteger;
              aMonthCount := aDataReader.FieldByName('BUNDLEMON').AsInteger;
            end;

          aAmtPrice := aMonthCount * aDataReader.FieldByName('MONTHAMT').AsInteger;
          MakeCell(aColIndex+2,aRowIndex,IntToStr(aAmtPrice),-1,xlRight);
          MakeCell(aColIndex+3,aRowIndex,IntToStr(aService),-1,xlRight);
          MakeCell(aColIndex+4,aRowIndex,IntToStr(aAmtPrice-aService),-1,xlRight);
          Inc(aRowIndex);
          aDataReader.Next
        end;
      //�N�S��ƪ����ɤW�~��
      if aRowIndex < aMaxRowIndex then
        for aIndex3 := 1 to (aMaxRowIndex - aRowIndex) do
          for aIndex2 :=0 to 4 do
            MakeCell(aColIndex+aIndex2,aRowIndex+aIndex3-1,'',-1,xlCenter);
      if aRowIndex > aMaxRowIndex then
        aMaxRowIndex := aRowIndex;
      aColIndex := aColIndex + 5;
    end;

  //�զX���~�W��
  MergeCell(1,4,1,aMaxRowIndex -1);
  MakeCell(1,4,aDescription,-1,xlLeft);
  
  //���`�p
  for aRowIndex := 4 to aMaxRowIndex -1 do
    begin
      aIndex2 := 4;
      aAmtPrice := 0;
      aService := 0;
      aTotal := 0;
      for aIndex := 0 to aList.Count-1 do
        begin
          aAmtPrice := aAmtPrice + ExcelApp.Cells[aRowIndex, aIndex * 5 + aIndex2].Value;
          aService := aService + ExcelApp.Cells[aRowIndex, aIndex * 5 + aIndex2 +1].Value;
          aTotal := aTotal + ExcelApp.Cells[aRowIndex, aIndex * 5 + aIndex2 +2].Value;
        end;
      MakeCell(aColIndex,aRowIndex,IntToStr(aAmtPrice),-1,xlRight);
      MakeCell(aColIndex+1,aRowIndex,IntToStr(aService),-1,xlRight);
      MakeCell(aColIndex+2,aRowIndex,IntToStr(aTotal),-1,xlRight);
    end;

  MakeGutter(aMaxRowIndex);
  aRowIndex := aMaxRowIndex + 1;
  //�����O�Ω��Y�C
  MakeCell(1,aRowIndex,'�����O��',34,xlCenter);
  aColIndex := 2;
  for aIndex := 0 to aList.Count-1 do
    begin
      MakeCell(aColIndex,aRowIndex,'���O����',34,xlLeft);
      MakeCell(aColIndex+1,aRowIndex,'ú�O'+#10+'����',34,xlCenter);
      MakeCell(aColIndex+2,aRowIndex,'�P��',34,xlCenter);
      MakeCell(aColIndex+3,aRowIndex,'�u�f��',34,xlCenter);
      MakeCell(aColIndex+4,aRowIndex,'���B'+#10+'�t��',34,xlCenter);
      aColIndex := aColIndex + 5;
    end;
  MakeCell(aColIndex,aRowIndex,'�P��'+#10+'�`�p',34,xlCenter);
  MakeCell(aColIndex+1,aRowIndex,'�u�f��'+#10+'�`�p',34,xlCenter);
  MakeCell(aColIndex+2,aRowIndex,'���B�t��'+#10+'�`�p',34,xlCenter);

  //�����O�ζ���
  aColIndex := 2;
//  aMaxRowIndex := -1;
  aInitRowIndex := aRowIndex + 1;
  for aIndex := 0 to aList.Count-1 do
    begin
      aRowIndex := aInitRowIndex;
      aDataReader.Close;
      aDataReader.Sql.Text := Format ( 'SELECT CITEMNAME, PERIOD1, MONTHAMT,  ' +
                                       'DISCOUNTAMT1, DEPOSITNAME, DEPOSITAMT ' +
                                       'FROM %s.CD078A                        ' +
                                       'WHERE BPCODE=''%s''                   ' +
                                       'AND SERVICETYPE=''%s''                ' ,
                                       [DataController.LoginInfo.DbAccount,
                                        BPCODE, aList[aIndex]] );
      aDataReader.Open;
      while not aDataReader.Eof do
        begin
          //�զX���~�W��
          MakeCell(1,aRowIndex,aDescription,-1,xlLeft);
          MakeCell(aColIndex,aRowIndex,aDataReader.FieldByName('CITEMNAME').AsString,-1,xlLeft);
          MakeCell(aColIndex+1,aRowIndex,aDataReader.FieldByName('PERIOD1').AsString,-1,xlRight);

          //�����P�����u�f���B
          aAmtPrice := aDataReader.FieldByName('PERIOD1').AsInteger *
                       aDataReader.FieldByName('MONTHAMT').AsInteger;
          aService := aDataReader.FieldByName('DISCOUNTAMT1').AsInteger;
          MakeCell(aColIndex+2,aRowIndex,IntToStr(aAmtPrice),-1,xlRight);
          MakeCell(aColIndex+3,aRowIndex,IntToStr(aService),-1,xlRight);
          MakeCell(aColIndex+4,aRowIndex,IntToStr(aAmtPrice-aService),-1,xlRight);

          //�O�Ҫ�
          MakeCell(1,aRowIndex+1,'�O�Ҫ�',-1,xlLeft);
          MakeCell(aColIndex,aRowIndex+1,aDataReader.FieldByName('DEPOSITNAME').AsString,-1,xlLeft);
          MakeCell(aColIndex+1,aRowIndex+1,'',-1,xlRight);
          MakeCell(aColIndex+2,aRowIndex+1,aDataReader.FieldByName('DEPOSITAMT').AsString,-1,xlRight);
          MakeCell(aColIndex+3,aRowIndex+1,aDataReader.FieldByName('DEPOSITAMT').AsString,-1,xlRight);
          MakeCell(aColIndex+4,aRowIndex+1,'0',-1,xlRight);

          aRowIndex := aRowIndex + 2;
          aDataReader.Next
        end;
      //�N�S��ƪ����ɤW�~��
      if aRowIndex < aMaxRowIndex then
        for aIndex3 := 1 to (aMaxRowIndex - aRowIndex) do
          for aIndex2 :=0 to 4 do
            MakeCell(aColIndex+aIndex2,aRowIndex+aIndex3-1,'',-1,xlCenter);
      if aRowIndex > aMaxRowIndex then
        aMaxRowIndex := aRowIndex;
      aColIndex := aColIndex + 5;
    end;

  //�����O���`�p
  for aRowIndex := aInitRowIndex to aMaxRowIndex -1 do
    begin
      aIndex2 := 4;
      aAmtPrice := 0;
      aService := 0;
      aTotal := 0;
      for aIndex := 0 to aList.Count-1 do
        begin
          aAmtPrice := aAmtPrice + ExcelApp.Cells[aRowIndex, aIndex * 5 + aIndex2].Value;
          aService := aService + ExcelApp.Cells[aRowIndex, aIndex * 5 + aIndex2 +1].Value;
          aTotal := aTotal + ExcelApp.Cells[aRowIndex, aIndex * 5 + aIndex2 +2].Value;
        end;
      MakeCell(aColIndex,aRowIndex,IntToStr(aAmtPrice),-1,xlRight);
      MakeCell(aColIndex+1,aRowIndex,IntToStr(aService),-1,xlRight);
      MakeCell(aColIndex+2,aRowIndex,IntToStr(aTotal),-1,xlRight);
    end;

  //�����O�ΦX�p
  MakeCell(1,aRowIndex,'�X�p',36,xlLeft);
  aColIndex := 2;
  for aIndex := 0 to aList.Count-1 do
    begin
      MakeCell(aColIndex,aRowIndex,'',36,xlLeft);
      MakeCell(aColIndex+1,aRowIndex,'',36,xlRight);
      aAmtPrice := 0;
      aService := 0;
      aTotal := 0;
      for aIndex2 := aInitRowIndex to aRowIndex -1 do
        begin
          aAmtPrice := aAmtPrice + ExcelApp.Cells[aIndex2, aColIndex+2].Value;
          aService := aService + ExcelApp.Cells[aIndex2, aColIndex+3].Value;
          aTotal := aTotal + ExcelApp.Cells[aIndex2, aColIndex+4].Value;
        end;
      MakeCell(aColIndex+2,aRowIndex,IntToStr(aAmtPrice),36,xlRight);
      MakeCell(aColIndex+3,aRowIndex,IntToStr(aService),36,xlRight);
      MakeCell(aColIndex+4,aRowIndex,IntToStr(aTotal),36,xlRight);
      aColIndex := aColIndex + 5;
    end;
  aAmtPrice := 0;
  aService := 0;
  aTotal := 0;
  for aIndex2 := aInitRowIndex to aRowIndex -1 do
    begin
      aAmtPrice := aAmtPrice + ExcelApp.Cells[aIndex2, aColIndex].Value;
      aService := aService + ExcelApp.Cells[aIndex2, aColIndex+1].Value;
      aTotal := aTotal + ExcelApp.Cells[aIndex2, aColIndex+2].Value;
    end;
  MakeCell(aColIndex,aRowIndex,IntToStr(aAmtPrice),36,xlRight);
  MakeCell(aColIndex+1,aRowIndex,IntToStr(aService),36,xlRight);
  MakeCell(aColIndex+2,aRowIndex,IntToStr(aTotal),36,xlRight);

  MakeGutter(aRowIndex+1);
  aRowIndex := aRowIndex + 2;
  //���u����Y�C
  MakeCell(1,aRowIndex,'���u��',34,xlCenter);
  aColIndex := 2;
  for aIndex := 0 to aList.Count-1 do
    begin
      MakeCell(aColIndex,aRowIndex,'���O����',34,xlLeft);
      MakeCell(aColIndex+1,aRowIndex,'',34,xlCenter);
      MakeCell(aColIndex+2,aRowIndex,'�P��',34,xlCenter);
      MakeCell(aColIndex+3,aRowIndex,'�u�f��',34,xlCenter);
      MakeCell(aColIndex+4,aRowIndex,'���B'+#10+'�t��',34,xlCenter);
      aColIndex := aColIndex + 5;
    end;
  MakeCell(aColIndex,aRowIndex,'�P��'+#10+'�`�p',34,xlCenter);
  MakeCell(aColIndex+1,aRowIndex,'�u�f��'+#10+'�`�p',34,xlCenter);
  MakeCell(aColIndex+2,aRowIndex,'���B�t��'+#10+'�`�p',34,xlCenter);

















  //��e�զ��̾A�j�p
  for aIndex := 1 to aColIndex+2 do
    ExcelApp.Columns[aIndex].EntireColumn.AutoFit;

  //�����A��Ц^�� A1
  ExcelApp.Cells[1,1].select;


{ ExcelApp.Cells.Select;
  ExcelApp.Cells.Font.Name := '�s�ө���';
  ExcelApp.Cells.Font.FontStyle := '�з�';
  ExcelApp.Cells.Font.Size := 10;

  MakeCell(1,3,'�զX���~',34,xlCenter);
  MakeGutter(5);
  MakeCell(1,6,'�����O��',34,xlCenter);
  MakeCell(1,8,'�X�p', 36,xlLeft);
  MakeGutter(9);
  MakeCell(1,10,'���u��',34,xlCenter);
  MakeCell(1,11,'�ѦҸ��A���u���O',34,xlCenter);

  MakeGutter(14);
  MakeCell(1,15,'�ѦҸ��A���u���O',34,xlCenter);
  MakeCell(1,17,'�X�p',36,xlLeft);
  ExcelApp.Columns[1].EntireColumn.AutoFit;

  for aIndex := 0 to AList.Count-1 do
    begin
      aServiceType := AList[aIndex];
      MergeCell(aIndex+2,1,aIndex+6,1);
      MakeCell(aIndex+2,1,aServiceType,34,xlCenter);


      MakeCell(aIndex+2,6,'���O����',34,xlLeft);
      MakeCell(aIndex+3,6,'ú�O'+#10+'����',34,xlCenter);
      MakeCell(aIndex+4,6,'�P��',34,xlCenter);
      MakeCell(aIndex+5,6,'�u�f��',34,xlCenter);
      MakeCell(aIndex+6,6,'���B'+#10+'�t��',34,xlCenter);

      MakeCell(aIndex+2,10,'���O����',34,xlLeft);
      MakeCell(aIndex+3,10,'',34,xlCenter);
      MakeCell(aIndex+4,10,'�P��',34,xlCenter);
      MakeCell(aIndex+5,10,'�u�f��',34,xlCenter);
      MakeCell(aIndex+6,10,'���B'+#10+'�t��',34,xlCenter);

      MergeCell(aIndex+2,11,aIndex+6,11);
      MakeCell(aIndex+2,11,'CATV�˾�',34,xlCenter);

      MergeCell(aIndex+2,15,aIndex+6,15);
      MakeCell(aIndex+2,15,'CATV�_��',34,xlCenter);

      ExcelApp.Columns[aIndex+2].EntireColumn.AutoFit;
      ExcelApp.Columns[aIndex+3].EntireColumn.AutoFit;
      ExcelApp.Columns[aIndex+4].EntireColumn.AutoFit;
      ExcelApp.Columns[aIndex+5].EntireColumn.AutoFit;
      ExcelApp.Columns[aIndex+6].EntireColumn.AutoFit;
    end;}

  //�~����
  //�M���ĤG�C���Ȧs���
//  aColIndex := 2;
//  for aIndex := 0 to aList.Count-1 do
//    begin
//      MakeCell(aColIndex,2,'',-1,xlCenter);
//      aColIndex := aColIndex + 5;
//    end;

  aDataReader.Close;
  aList.Free;
end;

procedure MakeCell(X, Y: Integer; ShowText: string; ColorIndex, AlignWhere: Integer);
begin
  ExcelApp.Cells[Y,X].Select;
  ExcelApp.ActiveCell.FormulaR1C1 := ShowText;
  ExcelApp.Selection.Interior.ColorIndex := ColorIndex;
  ExcelApp.Selection.HorizontalAlignment := AlignWhere;
  ExcelApp.Selection.VerticalAlignment := xlCenter;
  ExcelApp.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
  ExcelApp.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
  ExcelApp.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
  ExcelApp.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
end;

procedure MakeGutter(RowPos: Integer); 
begin
  ExcelApp.Rows[RowPos].Select;
  ExcelApp.Selection.RowHeight := ExcelApp.Selection.RowHeight div 2;
end;

procedure MergeCell(X1, Y1, X2, Y2: Integer);
begin
  ExcelApp.Range[ExcelApp.Cells[Y1,X1],ExcelApp.Cells[Y2, X2]].Select;
  ExcelApp.Selection.MergeCells := True;
end;

end.
