unit xmlU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, xmldom, XMLIntf, StdCtrls, Buttons, msxmldom, XMLDoc;

Type


  TMyXML = class
  private

    class procedure ParseStrings(Strs:string; Sep:Char; var I_StrList:TStringList) ; overload ;
  public
    class function createRootNode(I_XMLDocument : TXMLDocument; sI_NodeName, sI_NodeValue:String; var sI_Result:String):IXmlNode;
    class function addChild(I_ParentNode: IXmlNode; I_ChildNodeName, I_ChildNodeValue:String):IXmlNode;
    class procedure setAttribute(I_Node:IXmlNode; sI_AttributeName, sI_AttributeValue:STring);
    class function getNodeValue(I_Node:IXmlNode):String;
    class function getChildNodeValue(I_ParentNode:IXmlNode; sI_ChildNodeName:STring):String;    
    class function getAttributeValue(I_Node:IXmlNode; sI_AttributeName:String):String;
  end;


implementation

{ TMyXML }


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



class function TMyXML.addChild(I_ParentNode: IXmlNode; I_ChildNodeName,
  I_ChildNodeValue: String):IXmlNode;
var
    L_Result : IXmlNode;
begin
    L_Result := I_ParentNode.AddChild(I_ChildNodeName);
    L_Result.Text := I_ChildNodeValue;
    Result := L_Result;


end;

class procedure TMyXML.setAttribute(I_Node: IXmlNode; sI_AttributeName,
  sI_AttributeValue: STring);
begin
    I_Node.SetAttribute(sI_AttributeName,sI_AttributeValue);
end;

class function TMyXML.getNodeValue(I_Node: IXmlNode): String;
var
    sL_Result : String;
    vL_Result : OleVariant;

begin
    vL_Result := I_Node.getNodeValue;
    if vL_Result=null then
      sL_Result := ''
    else
      sL_Result := vL_Result;

    Result := sL_Result;

end;

class function TMyXML.getAttributeValue(I_Node: IXmlNode;
  sI_AttributeName:STring): String;
var
    sL_Result : String;
    vL_Result : OleVariant;
begin
    vL_Result := I_Node.GetAttribute(sI_AttributeName);
    if vL_Result=null then
      sL_Result := ''
    else
      sL_Result := vL_Result;

    Result := sL_Result;
end;

class function TMyXML.getChildNodeValue(I_ParentNode: IXmlNode;
  sI_ChildNodeName: STring): String;
var
    sL_Result : String;
    vL_Result : OleVariant;

begin
    vL_Result := I_ParentNode.GetChildValue(sI_ChildNodeName);
    if vL_Result=null then
      sL_Result := ''
    else
      sL_Result := vL_Result;

    Result := sL_Result;

end;

class function TMyXML.createRootNode(I_XMLDocument: TXMLDocument;
  sI_NodeName, sI_NodeValue: String; var sI_Result: String): IXmlNode;
begin
    sI_Result := '';
    if not Assigned(I_XMLDocument.DocumentElement) then
    begin
      I_XMLDocument.DocumentElement := I_XMLDocument.AddChild(sI_NodeName);
      I_XMLDocument.DocumentElement.Text := sI_NodeValue;
    end
    else
      sI_Result := '此 XML 已經有 Root Node 了';
    result := I_XMLDocument.DocumentElement;

end;

end.
