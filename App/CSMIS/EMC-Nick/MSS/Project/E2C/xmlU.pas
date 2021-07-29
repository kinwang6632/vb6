unit xmlU;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Db, Forms, Dialogs, Math,
  dbclient,ADOdb, Variants, MSXML_TLB ;

Type
  TIUDObj = class(TObject)
    sFieldName : String;
    sNewValue, sOldValue, sKey : String;
  end;


  TMyXML = class
  private
    class function GetFieldType(sI_FiledName:String):TFieldType;
    
    class procedure ParseStrings(Strs:string; Sep:Char; var I_StrList:TStringList) ; overload ;
  public
    class function getAttributeValue(I_Node:IXMLDOMNODE; sI_AttributeName : String):String;
    class function getSingleNodeText(I_ParentNode:IXMLDOMNODE; sI_NodeName : String):String;
    class procedure UpdateDataToDB(I_AppendAdo, I_adoQuery:TADOQuery; sI_TableName, sI_XML : String; bI_UseUpdateFun:Boolean);
    class procedure GetXMLData(I_Node : IXMLDOMNode;I_CDS:TClientDataSet);
    class function GetBottomNode(I_Node : IXMLDOMNode):IXMLDOMNode;
    class procedure GetNodesInfo(I_Nodes:IXMLDOMNodeList; var sI_NodesInfo:String);
    class procedure GetNodeInfo(I_Node:IXMLDOMNode; var sI_NodeName, sI_NodeText:String);    
    class function TransXMLFileToStr(I_XMLDoc:IXMLDOMDocument):String;overload;
    class function TransXMLFileToStr(I_XMLFileName:String):String;overload;
    class procedure CreateXMLFile(sI_XMLFileName:String;I_DS:TDataSet);
    class procedure AppendXMLDataFromTrascDS(I_XMLDoc : IXMLDOMDocument;I_DS:TDataSet;nI_UpdateKind:TUpdateKind; sG_KeyFieldsPos : String);
    class procedure GetCDSDataByXML(I_XMLDoc:IXMLDOMDocument; I_CDS:TClientDataSet; bI_EmptyDataSet:Boolean);overload;
    class procedure GetCDSDataByXML(I_XMLString:String; I_CDS:TClientDataSet; bI_EmptyDataSet:Boolean);overload;    
  end;


implementation



{ TMyXML }

class procedure TMyXML.AppendXMLDataFromTrascDS(I_XMLDoc : IXMLDOMDocument;I_DS: TDataSet;
  nI_UpdateKind: TUpdateKind; sG_KeyFieldsPos : String);
var
    L_RecTag, L_Node : IXMLDOMNode;
    sL_IUD : String;
    jj : Integer;
    sL_FieldName, sL_NewValue, sL_OldValue : String;
begin
    case nI_UpdateKind of
      ukModify:
       sL_IUD := 'U';
      ukInsert:
       sL_IUD := 'I';
      ukDelete:
       sL_IUD := 'D';
    end;
    L_RecTag := I_XMLDoc.createNode(NODE_ELEMENT,sL_IUD,'');
    I_XMLDoc.documentElement.appendChild(L_RecTag);

    with I_DS do
    begin
        for jj:=0 to Fields.Count - 1 do
        begin
          sL_FieldName := Fields[jj].FieldName;
          if (sL_IUD='I') then//Insert
          begin
            if (VarType(Fields[jj].NewValue)<>varEmpty)and(Fields[jj].NewValue<>NULL) then
            begin
              sL_NewValue := Fields[jj].NewValue;
              sL_OldValue := sL_NewValue;
            end
            else
            begin
              sL_NewValue := '';
              sL_OldValue := '';
            end;
          end
          else if (sL_IUD='D') then//Delete
          begin
            if (VarType(Fields[jj].NewValue)<>varEmpty) then
            begin
              if Fields[jj].NewValue<>NULL then
              begin
                sL_OldValue := Fields[jj].OldValue;
                sL_NewValue := sL_OldValue;
              end
              else
              begin
                sL_NewValue := '';
                sL_OldValue := '';
              end;
            end
            else
            begin
              sL_NewValue := '';
              if Fields[jj].OldValue<>NULL then
                sL_OldValue := Fields[jj].OldValue
              else
                sL_OldValue := '';
            end;
          end
          else//Update
          begin
            if Fields[jj].OldValue=NULL then
            begin
              if (VarType(Fields[jj].NewValue)<>varEmpty) then
              begin
                if Fields[jj].NewValue<>NULL then
                begin
                  sL_NewValue := Fields[jj].NewValue;
                  sL_OldValue := sL_NewValue;
                end
                else
                begin
                  sL_NewValue := '';
                  sL_OldValue := '';
                end;
              end
              else
              begin
                sL_NewValue := '';
                sL_OldValue := '';
              end;
            end
            else
            begin
              sL_OldValue := Fields[jj].OldValue;
              if (VarType(Fields[jj].NewValue)<>varEmpty) then
              begin
                if Fields[jj].NewValue<>NULL then
                  sL_NewValue := Fields[jj].NewValue
                else
                  sL_NewValue := '';
              end
              else
                sL_NewValue := sL_OldValue;
            end;
          end;
          {
          if (VarType(Fields[jj].NewValue)<>varEmpty) then
          begin
          if (Fields[jj].NewValue=NULL) then
            sL_NewValue := ''
          else
            sL_NewValue := Fields[jj].NewValue;
          end
          else
            sL_NewValue := '';

          if (Fields[jj].OldValue=NULL)then
            sL_OldValue := ''
          else
            sL_OldValue := Fields[jj].OldValue;

          if (sL_IUD='I') or (sL_IUD='D')then
          begin
            if sL_NewValue='' then
              sL_NewValue := sL_OldValue;
            if sL_OldValue='' then
              sL_OldValue := sL_NewValue;
          end
          else
          begin
            if (VarType(Fields[jj].NewValue)<>varEmpty) then
              if Fields[jj].NewValue=Fields[jj].OldValue then
               sL_NewValue := sL_OldValue;
          end;
          }
          L_Node:= I_XMLDoc.createNode(NODE_ELEMENT,sL_FieldName,'');
          if AnsiPos(IntToStr(jj),sG_KeyFieldsPos) <> 0 then
            L_Node.text := sL_OldValue+';'+sL_NewValue + ';' +  'KF'
          else
            L_Node.text := sL_OldValue+';'+sL_NewValue;



          L_RecTag.appendChild(L_Node);
        end;
    end;

end;

class procedure TMyXML.CreateXMLFile(sI_XMLFileName: String; I_DS: TDataSet);
var
    L_XMLDoc : IXMLDOMDocument;
    L_RecTag, L_Node : IXMLDOMNode;
    L_Element : IXMLDOMElement;
    ii,jj : Integer;
    sL_FieldName, sL_FieldValue : String;
begin
    L_XMLDoc := CoDOMDocument.Create;

    L_Element := L_XMLDoc.createElement(sI_XMLFileName);
//    L_XMLDoc.documentElement := L_Element;
    ii := 0 ;
    with I_DS do
    begin
      if State = dsInactive then
        Active := True;
      First;
      while not Eof do
      begin
        INC(ii);
        L_RecTag := L_XMLDoc.createNode(NODE_ELEMENT,'R'+IntToStr(ii),'');
        L_XMLDoc.documentElement.appendChild(L_RecTag);

        for jj:=0 to Fields.Count - 1 do
        begin
          if Fields[jj].FieldKind = fkData then
          begin
            sL_FieldName := Fields[jj].FieldName;
            sL_FieldValue := Fields[jj].AsString;

            L_Node:= L_XMLDoc.createNode(NODE_ELEMENT,sL_FieldName,'');
            L_Node.text := sL_FieldValue;

            L_RecTag.appendChild(L_Node);
          end;
        end;
        Next;
      end;
    end;
    L_XMLDoc.save(sI_XMLFileName+'.xml');
//    L_XMLDoc.Free;
    L_XMLDoc := nil;
end;

class function TMyXML.GetBottomNode(I_Node: IXMLDOMNode): IXMLDOMNode;
begin
    if I_Node=nil then Exit;
    while I_Node.hasChildNodes do
    begin
      I_Node := I_Node.firstChild;
    end;
    Result := I_Node.parentNode;
end;

class procedure TMyXML.GetCDSDataByXML(I_XMLDoc: IXMLDOMDocument;
  I_CDS: TClientDataSet; bI_EmptyDataSet:Boolean);
var
    L_Node : IXMLDOMNode;
begin
    if I_CDS.state = dsInactive then
      I_CDS.CreateDataSet;
    if bI_EmptyDataSet then
      I_CDS.EmptyDataSet;
    L_Node := GetBottomNode(I_XMLDoc.documentElement.firstChild);
    if L_Node <> nil then
    begin
      L_Node := L_Node.parentNode;
      while L_Node<>nil do
      begin
        GetXMLData(L_Node, I_CDS);
        L_Node := L_Node.nextSibling;
      end;
    end;
end;

class procedure TMyXML.GetCDSDataByXML(I_XMLString: String;
  I_CDS: TClientDataSet; bI_EmptyDataSet:Boolean);
var
    L_XMLDoc : IXMLDOMDocument;
begin
    L_XMLDoc := CoDOMDocument.Create;
    L_XMLDoc.loadXML(I_XMLString);
    {
    ReadyState :
     0:未載入
     1:載入中
     2:已載入
     3:唯讀,且僅顯示部分資料
     4:可讀寫並已可顯示完整資料
    }
    while L_XMLDoc.ReadyState<>4 do
      Sleep(1000);
    
    GetCDSDataByXML(L_XMLDoc,I_CDS, bI_EmptyDataSet);
    L_XMLDoc := nil;
end;

class procedure TMyXML.GetNodeInfo(I_Node: IXMLDOMNode; var sI_NodeName,
  sI_NodeText: String);
begin
    sI_NodeName := I_Node.nodeName;
    sI_NodeText := I_Node.text;
end;

class procedure TMyXML.GetNodesInfo(I_Nodes: IXMLDOMNodeList;
  var sI_NodesInfo: String);
var
    ii : Integer;
    sL_NodeValue, sL_RecName: String;
    L_Node : IXMLDOMNode;
begin
      for ii:=0 to I_Nodes.length -1 do
      begin
        L_Node := I_Nodes.item[ii];
        if L_Node.hasChildNodes then
        begin
          sL_RecName := I_Nodes.item[ii].nodeName;
          sI_NodesInfo := sI_NodesInfo + '<' +  sL_RecName + '>';
          GetNodesInfo(L_Node.childNodes, sI_NodesInfo);
          sI_NodesInfo := sI_NodesInfo + '</' + sL_RecName + '>';
        end;
        if L_Node.nodeType = NODE_TEXT Then
        begin
          sL_NodeValue := I_Nodes.item[ii].Text;
          sI_NodesInfo := sI_NodesInfo + sL_NodeValue;
        end;    

      end;
end;

class procedure TMyXML.GetXMLData(I_Node: IXMLDOMNode;
  I_CDS: TClientDataSet);
var
    ii : Integer;
    sL_FieldName, sL_FieldValue : String;
begin
    if I_Node.childNodes.length >0 then
      I_CDS.append;
    for ii:=0 to I_Node.childNodes.length -1 do
    begin
      sL_FieldName := I_Node.childNodes.item[ii].nodeName;
      sL_FieldValue := I_Node.childNodes.item[ii].text;
      I_CDS.FieldByName(sL_FieldName).AsString := sL_FieldValue;
    end;
    if I_CDS.State = dsInsert then
      I_CDS.post;
end;

class function TMyXML.TransXMLFileToStr(I_XMLDoc: IXMLDOMDocument): String;
var
    sL_Result : String;
    sL_TableName : String;
begin
    sL_TableName := I_XMLDoc.documentElement.nodeName;
    sL_Result := '<' + sL_TableName +'>';
    GetNodesInfo(I_XMLDoc.documentElement.childNodes, sL_Result);
    sL_Result := sL_Result + '</' + sL_TableName +'>';
    Result := sL_Result;
end;

class procedure TMyXML.ParseStrings(Strs: string; Sep: Char;
  var I_StrList: TStringList);
var
  ii, jj : Integer ;
begin
  I_StrList.Clear;
  if Strs = '' then Exit;
  jj := 1 ;
  for ii:= 1 to Length(Strs) do
  begin
    if Strs[ii] = Sep then
    begin
      I_StrList.Add(Trim(Copy(Strs, jj, ii-jj))) ;
      jj := ii + 1 ;
    end ;
  end ;
  I_StrList.Add(Trim(Copy(Strs, jj, Length(Strs)-jj+1))) ;
end;


class function TMyXML.TransXMLFileToStr(I_XMLFileName: String): String;
var
    L_XMLDoc : IXMLDOMDocument;
begin
    L_XMLDoc := CoDOMDocument.Create;
    L_XMLDoc.load(I_XMLFileName);
    {
    ReadyState :
     0:未載入
     1:載入中
     2:已載入
     3:唯讀,且僅顯示部分資料
     4:可讀寫並已可顯示完整資料
    }
    while L_XMLDoc.ReadyState<>4 do
      Sleep(1000);
    Result := TransXMLFileToStr(L_XMLDoc);
    L_XMLDoc := nil;
end;

class procedure TMyXML.UpdateDataToDB(I_AppendAdo, I_adoQuery:TADOQuery; sI_TableName, sI_XML: String; bI_UseUpdateFun:Boolean);
var
    sL_NodeName, sL_NodeText : String;
    L_XMLDoc : IXMLDOMDocument;
    sL_IUD : String;
    ii, jj: Integer;
    L_PNode, L_CNode : IXMLDOMNode;
    sL_FieldName : String;
    L_StrList, L_IUD : TStringList;
    sL_NewValue, sL_OldValue, sL_Key : String;
    L_IUDObj : TIUDObj;

    procedure UpdateFun(I_IUDStrList:TStringList);
    var
        sL_Flag : String;
        sL_ParamName : String;
        ii, nL_Ndx : Integer;
        sL_FieldName : String;
        sL_WhereFieldsName, sL_ValueFieldsName : String;
    begin
        with I_adoQuery do
        begin
          Close;
          SQL.Clear;
          SQL.Add('UPDATE ' + sI_TableName );
          for ii:=0 to I_IUDStrList.Count-1 do
          begin
            if (I_IUDStrList.Objects[ii] As TIUDObj).sKey = 'N' then
            begin
            sL_FieldName := (I_IUDStrList.Objects[ii] As TIUDObj).sFieldName;
            if Length(sL_ValueFieldsName)=0 then
              sL_ValueFieldsName := sL_FieldName + '=:' + sL_FieldName + 'V'
            else
              sL_ValueFieldsName := sL_ValueFieldsName +  ',' + sL_FieldName + '=:' + sL_FieldName + 'V';
            end;
          end;
          SQL.Add('SET '+sL_ValueFieldsName);

          for ii:=0 to I_IUDStrList.Count-1 do
          begin
            if (I_IUDStrList.Objects[ii] As TIUDObj).sKey = 'Y' then
            begin
              sL_FieldName := (I_IUDStrList.Objects[ii] As TIUDObj).sFieldName ;
              if Length(sL_WhereFieldsName)=0 then
                sL_WhereFieldsName := sL_FieldName + '=:' + sL_FieldName+ 'W'
              else
                sL_WhereFieldsName := sL_WhereFieldsName + ' and ' + sL_FieldName + '=:' + sL_FieldName+ 'W';
            end;
          end;
          SQL.Add('WHERE '+sL_WhereFieldsName);
          for ii:=0 to Parameters.Count -1 do
          begin

            sL_Flag := Copy(Parameters[ii].Name, Length(Parameters[ii].Name),1);
            sL_ParamName := Copy(Parameters[ii].Name, 1, Length(Parameters[ii].Name)-1);
            nL_Ndx := I_IUDStrList.IndexOf(sL_ParamName);
            if nL_Ndx<> -1 then
            begin
              if sL_Flag='V' then//Values...
              begin
                if (I_IUDStrList.Objects[nL_Ndx] As TIUDObj).sNewValue='' then
                begin
                  Parameters[ii].DataType := GetFieldType((I_IUDStrList.Objects[ii] As TIUDObj).sFieldName);
                  Parameters[ii].Value := NULL;
                end
                else
                  Parameters[ii].Value := (I_IUDStrList.Objects[nL_Ndx] As TIUDObj).sNewValue;
              end
              else//"W" => Where...
              begin
                if (I_IUDStrList.Objects[nL_Ndx] As TIUDObj).sOldValue='' then
                begin
                  Parameters[ii].DataType := GetFieldType((I_IUDStrList.Objects[ii] As TIUDObj).sFieldName);
                  Parameters[ii].Value := NULL;
                end
                else
                  Parameters[ii].Value := (I_IUDStrList.Objects[nL_Ndx] As TIUDObj).sOldValue;
              end;
            end;
          end;
          ExecSQL;
          Close;
        end;
    end;

    procedure DeleteFun(I_IUDStrList:TStringList);
    var
        ii, nL_Ndx : Integer;
        sL_FieldName : String;
        sL_WhereFieldsName : String;
    begin
        with I_adoQuery do
        begin
          Close;
          SQL.Clear;
          SQL.Add('delete ' + sI_TableName );
          for ii:=0 to I_IUDStrList.Count-1 do
          begin
            if (I_IUDStrList.Objects[ii] As TIUDObj).sKey = 'Y' then
            begin
              sL_FieldName := (I_IUDStrList.Objects[ii] As TIUDObj).sFieldName;
              if Length(sL_WhereFieldsName)=0 then
                sL_WhereFieldsName := sL_FieldName + '=:' + sL_FieldName
              else
                sL_WhereFieldsName := sL_WhereFieldsName + ' and ' + sL_FieldName + '=:' + sL_FieldName;
            end;
          end;
          SQL.Add('WHERE '+sL_WhereFieldsName);

          for ii:=0 to Parameters.Count -1 do
          begin
            nL_Ndx := I_IUDStrList.IndexOf(Parameters[ii].Name);
            if nL_Ndx<> -1 then
            begin
              if (I_IUDStrList.Objects[nL_Ndx] As TIUDObj).sNewValue='' then
              begin
                Parameters[ii].DataType := GetFieldType((I_IUDStrList.Objects[ii] As TIUDObj).sFieldName);
                Parameters[ii].Value := NULL;
              end
              else
                Parameters[ii].Value := (I_IUDStrList.Objects[nL_Ndx] As TIUDObj).sNewValue;
            end;
          end;
          ExecSQL;
          Close;
        end;
    end;

    procedure InsertFun(I_IUDStrList:TStringList);
    var
        sL_Flag : String;
        sL_ParamName : String;
        ii, nL_Ndx : Integer;
        sL_FieldName : String;
        sL_WhereFieldsName, sL_ValueFieldsName : String;
    begin
        with I_AppendAdo do
        begin
          if state= dsInactive then
            Active := True;
          Append;
          for ii:=0 to I_IUDStrList.Count-1 do
          begin
            if (I_IUDStrList.Objects[ii] As TIUDObj).sNewValue='' then
            begin
              FieldByName((I_IUDStrList.Objects[ii] As TIUDObj).sFieldName).Value := NULL;
            end
            else
            begin
              FieldByName((I_IUDStrList.Objects[ii] As TIUDObj).sFieldName).Value :=(I_IUDStrList.Objects[ii] As TIUDObj).sNewValue;
            end;
          end;
          Post;
        end;
        Exit;


        //down,因為delphi 5似乎對用param的方式不support,故先mark掉,改用append
        {
        sL_ValueFieldsName := '(';
        with I_adoQuery do
        begin
          Close;
          SQL.Clear;
          SQL.Add('Insert ' + sI_TableName );
          for ii:=0 to I_IUDStrList.Count-1 do
          begin
            if ii=0 then
              sL_ValueFieldsName := sL_ValueFieldsName + ':' + (I_IUDStrList.Objects[ii] As TIUDObj).sFieldName
            else
              sL_ValueFieldsName := sL_ValueFieldsName +  ',:' + (I_IUDStrList.Objects[ii] As TIUDObj).sFieldName;
          end;
          sL_ValueFieldsName := sL_ValueFieldsName + ')';
          SQL.Add('VALUES'+sL_ValueFieldsName);

          for ii:=0 to I_IUDStrList.Count-1 do
          begin
            if (I_IUDStrList.Objects[ii] As TIUDObj).sNewValue='' then
            begin
              Parameters[ii].DataType := GetFieldType((I_IUDStrList.Objects[ii] As TIUDObj).sFieldName);
              Parameters[ii].Value := NULL;
            end
            else
              Parameters[ii].Value := (I_IUDStrList.Objects[ii] As TIUDObj).sNewValue;

          end;
          ExecSQL;
          Close;
        end;
        }
        //up,因為delphi 5似乎對用param的方式不support,故先mark掉,改用append
    end;

    procedure UpdateToDB(sI_IUD:String;I_IUDStrList:TStringList);
    begin
      if sI_IUD = 'I' then//Insert...
      begin
        DeleteFun(I_IUDStrList);
        InsertFun(I_IUDStrList);
      end
      else if sI_IUD = 'U' then//Update...
      begin
        if bI_UseUpdateFun then
          UpdateFun(I_IUDStrList)
        else
        begin
          DeleteFun(I_IUDStrList);
          InsertFun(I_IUDStrList);
        end;
      end
      else if sI_IUD = 'D' then//Delete...
      begin
        DeleteFun(I_IUDStrList);
      end;
    end;

begin

    L_StrList := TStringList.Create;
    L_IUD := TStringList.Create;
    L_XMLDoc := CoDOMDocument.Create;
    L_XMLDoc.loadXML(sI_XML);


    
    {
    ReadyState :
     0:未載入
     1:載入中
     2:已載入
     3:唯讀,且僅顯示部分資料
     4:可讀寫並已可顯示完整資料
    }
    while L_XMLDoc.ReadyState<>4 do
      Sleep(1000);

    for ii:=0 to L_XMLDoc.documentElement.childNodes.length -1 do
    begin
      L_PNode := L_XMLDoc.documentElement.childNodes.item[ii];
      sL_IUD := L_PNode.nodeName;

      for jj:=0 to L_PNode.childNodes.length -1 do
      begin
        L_CNode := L_PNode.childNodes.item[jj];
        GetNodeInfo(L_CNode,sL_NodeName, sL_NodeText);
        sL_FieldName := sL_NodeName;
        ParseStrings(sL_NodeText,';',L_StrList);
        sL_OldValue := L_StrList.Strings[0];
        sL_NewValue := L_StrList.Strings[1];

        if L_StrList.Count=3 then//為key
          sL_Key := 'Y'
        else
          sL_Key := 'N';
        L_IUDObj := TIUDObj.Create;

        L_IUDObj.sFieldName := sL_FieldName;
        L_IUDObj.sNewValue := sL_NewValue;
        L_IUDObj.sOldValue := sL_OldValue;
        L_IUDObj.sKey := sL_Key;
        L_IUD.AddObject(sL_FieldName, L_IUDObj);
      end;

      UpdateToDB(sL_IUD, L_IUD);
      L_IUD.Clear;

    end;

    L_XMLDoc := nil;
    L_StrList.Free;
    L_IUD.Free;

end;

class function TMyXML.GetFieldType(sI_FiledName: String): TFieldType;
var
    sL_FieldType : String;
    L_FT : TFieldType;
begin
    sL_FieldType := UpperCase(Copy(sI_FiledName,1,1));
    if sL_FieldType='N' then//Integer
      L_FT := ftInteger
    else if sL_FieldType='F' then//Float
      L_FT := ftFloat
    else if sL_FieldType='D' then//Date
      L_FT := ftDateTime
    else if sL_FieldType='B' then//Boolean
      L_FT := ftBoolean
    else if sL_FieldType='S' then//String
      L_FT := ftString
    else
      L_FT := ftUnknown;

    Result := L_FT;  
end;

class function TMyXML.getSingleNodeText(I_ParentNode: IXMLDOMNODE;
  sI_NodeName: String): String;
var
    L_TmpNode : IXMLDOMNODE;
    sL_NodeName, sL_NodeText : String;
begin
        L_TmpNode := I_ParentNode.selectSingleNode('./'+sI_NodeName); //表示從現行 node 開始 search
        TMyXML.GetNodeInfo(L_TmpNode, sL_NodeName, sL_NodeText);
        result := sL_NodeText; 

end;

class function TMyXML.getAttributeValue(I_Node: IXMLDOMNODE;
  sI_AttributeName: String): String;
begin
    result := I_Node.attributes.getNamedItem(sI_AttributeName).text;
end;

end.
