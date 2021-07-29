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

  //先找出 BPCODE 有哪些服務類別
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

  //第一行
  MakeCell(1,1,'服務類別',34,xlCenter);
  aColIndex := 2;
  for aIndex := 0 to aList.Count-1 do
    begin
      MergeCell(aColIndex,1,aColIndex+4,1);
      MakeCell(aColIndex,1,aList[aIndex],34,xlCenter);
      //第二行，用來記錄每個 ServideType 有幾個項目，等全部匯出到 EXCEL 再清除
//      MergeCell(aColIndex,2,aColIndex+4,2);
//      MakeCell(aColIndex,2,'0',-1,xlCenter);
      aColIndex := aColIndex + 5;
    end;
  MakeCell(aColIndex,1,'牌價'+#10+'總計',34,xlCenter);
  MakeCell(aColIndex+1,1,'優惠價'+#10+'總計',34,xlCenter);
  MakeCell(aColIndex+2,1,'金額差異'+#10+'總計',34,xlCenter);

  MakeGutter(2);

  //組合產品抬頭列
  MakeCell(1,3,'組合產品',34,xlCenter);
  aColIndex := 2;
  for aIndex := 0 to aList.Count-1 do
    begin
      MakeCell(aColIndex,3,'收費項目',34,xlLeft);
      MakeCell(aColIndex+1,3,'綁約'+#10+'月數',34,xlCenter);
      MakeCell(aColIndex+2,3,'牌價',34,xlCenter);
      MakeCell(aColIndex+3,3,'優惠價',34,xlCenter);
      MakeCell(aColIndex+4,3,'金額'+#10+'差異',34,xlCenter);
      aColIndex := aColIndex + 5;
    end;
  MakeCell(aColIndex,3,'牌價'+#10+'總計',34,xlCenter);
  MakeCell(aColIndex+1,3,'優惠價'+#10+'總計',34,xlCenter);
  MakeCell(aColIndex+2,3,'金額差異'+#10+'總計',34,xlCenter);

  //組合產品項目
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

          //階段式優惠月數及優惠金額合計
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
              //綁約月數 <> 0
              //若優惠月數小於綁約月數，要 (綁約月數-優惠月數) * 每月牌價
              //才是真正的優惠金額
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
      //將沒資料的欄位補上外框
      if aRowIndex < aMaxRowIndex then
        for aIndex3 := 1 to (aMaxRowIndex - aRowIndex) do
          for aIndex2 :=0 to 4 do
            MakeCell(aColIndex+aIndex2,aRowIndex+aIndex3-1,'',-1,xlCenter);
      if aRowIndex > aMaxRowIndex then
        aMaxRowIndex := aRowIndex;
      aColIndex := aColIndex + 5;
    end;

  //組合產品名稱
  MergeCell(1,4,1,aMaxRowIndex -1);
  MakeCell(1,4,aDescription,-1,xlLeft);
  
  //算總計
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
  //首期費用抬頭列
  MakeCell(1,aRowIndex,'首期費用',34,xlCenter);
  aColIndex := 2;
  for aIndex := 0 to aList.Count-1 do
    begin
      MakeCell(aColIndex,aRowIndex,'收費項目',34,xlLeft);
      MakeCell(aColIndex+1,aRowIndex,'繳費'+#10+'期數',34,xlCenter);
      MakeCell(aColIndex+2,aRowIndex,'牌價',34,xlCenter);
      MakeCell(aColIndex+3,aRowIndex,'優惠價',34,xlCenter);
      MakeCell(aColIndex+4,aRowIndex,'金額'+#10+'差異',34,xlCenter);
      aColIndex := aColIndex + 5;
    end;
  MakeCell(aColIndex,aRowIndex,'牌價'+#10+'總計',34,xlCenter);
  MakeCell(aColIndex+1,aRowIndex,'優惠價'+#10+'總計',34,xlCenter);
  MakeCell(aColIndex+2,aRowIndex,'金額差異'+#10+'總計',34,xlCenter);

  //首期費用項目
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
          //組合產品名稱
          MakeCell(1,aRowIndex,aDescription,-1,xlLeft);
          MakeCell(aColIndex,aRowIndex,aDataReader.FieldByName('CITEMNAME').AsString,-1,xlLeft);
          MakeCell(aColIndex+1,aRowIndex,aDataReader.FieldByName('PERIOD1').AsString,-1,xlRight);

          //首期牌價及優惠金額
          aAmtPrice := aDataReader.FieldByName('PERIOD1').AsInteger *
                       aDataReader.FieldByName('MONTHAMT').AsInteger;
          aService := aDataReader.FieldByName('DISCOUNTAMT1').AsInteger;
          MakeCell(aColIndex+2,aRowIndex,IntToStr(aAmtPrice),-1,xlRight);
          MakeCell(aColIndex+3,aRowIndex,IntToStr(aService),-1,xlRight);
          MakeCell(aColIndex+4,aRowIndex,IntToStr(aAmtPrice-aService),-1,xlRight);

          //保證金
          MakeCell(1,aRowIndex+1,'保證金',-1,xlLeft);
          MakeCell(aColIndex,aRowIndex+1,aDataReader.FieldByName('DEPOSITNAME').AsString,-1,xlLeft);
          MakeCell(aColIndex+1,aRowIndex+1,'',-1,xlRight);
          MakeCell(aColIndex+2,aRowIndex+1,aDataReader.FieldByName('DEPOSITAMT').AsString,-1,xlRight);
          MakeCell(aColIndex+3,aRowIndex+1,aDataReader.FieldByName('DEPOSITAMT').AsString,-1,xlRight);
          MakeCell(aColIndex+4,aRowIndex+1,'0',-1,xlRight);

          aRowIndex := aRowIndex + 2;
          aDataReader.Next
        end;
      //將沒資料的欄位補上外框
      if aRowIndex < aMaxRowIndex then
        for aIndex3 := 1 to (aMaxRowIndex - aRowIndex) do
          for aIndex2 :=0 to 4 do
            MakeCell(aColIndex+aIndex2,aRowIndex+aIndex3-1,'',-1,xlCenter);
      if aRowIndex > aMaxRowIndex then
        aMaxRowIndex := aRowIndex;
      aColIndex := aColIndex + 5;
    end;

  //首期費用總計
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

  //首期費用合計
  MakeCell(1,aRowIndex,'合計',36,xlLeft);
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
  //派工單抬頭列
  MakeCell(1,aRowIndex,'派工單',34,xlCenter);
  aColIndex := 2;
  for aIndex := 0 to aList.Count-1 do
    begin
      MakeCell(aColIndex,aRowIndex,'收費項目',34,xlLeft);
      MakeCell(aColIndex+1,aRowIndex,'',34,xlCenter);
      MakeCell(aColIndex+2,aRowIndex,'牌價',34,xlCenter);
      MakeCell(aColIndex+3,aRowIndex,'優惠價',34,xlCenter);
      MakeCell(aColIndex+4,aRowIndex,'金額'+#10+'差異',34,xlCenter);
      aColIndex := aColIndex + 5;
    end;
  MakeCell(aColIndex,aRowIndex,'牌價'+#10+'總計',34,xlCenter);
  MakeCell(aColIndex+1,aRowIndex,'優惠價'+#10+'總計',34,xlCenter);
  MakeCell(aColIndex+2,aRowIndex,'金額差異'+#10+'總計',34,xlCenter);

















  //欄寬調成最適大小
  for aIndex := 1 to aColIndex+2 do
    ExcelApp.Columns[aIndex].EntireColumn.AutoFit;

  //完成，游標回到 A1
  ExcelApp.Cells[1,1].select;


{ ExcelApp.Cells.Select;
  ExcelApp.Cells.Font.Name := '新細明體';
  ExcelApp.Cells.Font.FontStyle := '標準';
  ExcelApp.Cells.Font.Size := 10;

  MakeCell(1,3,'組合產品',34,xlCenter);
  MakeGutter(5);
  MakeCell(1,6,'首期費用',34,xlCenter);
  MakeCell(1,8,'合計', 36,xlLeft);
  MakeGutter(9);
  MakeCell(1,10,'派工單',34,xlCenter);
  MakeCell(1,11,'參考號∕派工類別',34,xlCenter);

  MakeGutter(14);
  MakeCell(1,15,'參考號∕派工類別',34,xlCenter);
  MakeCell(1,17,'合計',36,xlLeft);
  ExcelApp.Columns[1].EntireColumn.AutoFit;

  for aIndex := 0 to AList.Count-1 do
    begin
      aServiceType := AList[aIndex];
      MergeCell(aIndex+2,1,aIndex+6,1);
      MakeCell(aIndex+2,1,aServiceType,34,xlCenter);


      MakeCell(aIndex+2,6,'收費項目',34,xlLeft);
      MakeCell(aIndex+3,6,'繳費'+#10+'期數',34,xlCenter);
      MakeCell(aIndex+4,6,'牌價',34,xlCenter);
      MakeCell(aIndex+5,6,'優惠價',34,xlCenter);
      MakeCell(aIndex+6,6,'金額'+#10+'差異',34,xlCenter);

      MakeCell(aIndex+2,10,'收費項目',34,xlLeft);
      MakeCell(aIndex+3,10,'',34,xlCenter);
      MakeCell(aIndex+4,10,'牌價',34,xlCenter);
      MakeCell(aIndex+5,10,'優惠價',34,xlCenter);
      MakeCell(aIndex+6,10,'金額'+#10+'差異',34,xlCenter);

      MergeCell(aIndex+2,11,aIndex+6,11);
      MakeCell(aIndex+2,11,'CATV裝機',34,xlCenter);

      MergeCell(aIndex+2,15,aIndex+6,15);
      MakeCell(aIndex+2,15,'CATV復機',34,xlCenter);

      ExcelApp.Columns[aIndex+2].EntireColumn.AutoFit;
      ExcelApp.Columns[aIndex+3].EntireColumn.AutoFit;
      ExcelApp.Columns[aIndex+4].EntireColumn.AutoFit;
      ExcelApp.Columns[aIndex+5].EntireColumn.AutoFit;
      ExcelApp.Columns[aIndex+6].EntireColumn.AutoFit;
    end;}

  //洗阿給
  //清掉第二列的暫存資料
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
