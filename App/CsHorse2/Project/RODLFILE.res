        ??  ??                  /   0   ??
 R O D L F I L E                     <?xml version="1.0" encoding="utf-8"?>
<Library Name="CsHorse2Library" UID="{2AD54CDB-1DA6-4043-8B23-FFBFDAC74174}" Version="3.0">
<Services>
<Service Name="LoginService" UID="{42B89B03-B702-4318-AF5D-244FDB7A476A}">
<Interfaces>
<Interface Name="Default" UID="{D4BF3D0A-AE86-4A14-A9F9-BC7592BEA6DD}">
<Operations>
<Operation Name="Login" UID="{72507EB7-A38E-484D-BE69-A7FED3E67709}">
<Parameters>
<Parameter Name="Result" DataType="Boolean" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="InOut" >
</Parameter>
<Parameter Name="AErrMsg" DataType="String" Flag="InOut" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="Logout" UID="{082A6523-990A-4CB3-81F2-3E33D5A612C5}">
<Parameters>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="GetClientParam" UID="{F9F4AE7C-D845-459D-839C-44D72BA0373A}">
<Parameters>
<Parameter Name="Result" DataType="Binary" Flag="Result">
</Parameter>
</Parameters>
</Operation>
<Operation Name="GetOraSysDate" UID="{2E562452-60A7-4273-96B7-EE41D2C37C5D}">
<Parameters>
<Parameter Name="Result" DataType="String" Flag="Result">
</Parameter>
<Parameter Name="ACompCode" DataType="String" Flag="In" >
</Parameter>
</Parameters>
</Operation>
</Operations>
</Interface>
</Interfaces>
</Service>
<Service Name="AnnService" UID="{1EC35274-5020-4935-BF27-0176D5D79FD2}">
<Interfaces>
<Interface Name="Default" UID="{B1A16ABD-12C1-47A8-9E47-1D107EEBC2DB}">
<Operations>
<Operation Name="GetSO021" UID="{85E4FC20-D59B-427C-9E70-65D35519F4DC}">
<Parameters>
<Parameter Name="Result" DataType="Binary" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="GetSOListText" UID="{64B53352-413F-4A3B-83A0-E5C7E88772C9}">
<Parameters>
<Parameter Name="Result" DataType="String" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="InOut" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="GetCD042" UID="{87CEE27A-3B50-4DA6-8F71-E59EE6154310}">
<Parameters>
<Parameter Name="Result" DataType="Binary" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="GetSO022" UID="{B6CD2AD9-01C9-41F1-8AA1-CEB1D40C9C1F}">
<Parameters>
<Parameter Name="Result" DataType="Binary" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="GetSO023" UID="{90FA26AC-584A-4210-840C-5CD0F07997EC}">
<Parameters>
<Parameter Name="Result" DataType="Binary" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
</Parameters>
</Operation>
</Operations>
</Interface>
</Interfaces>
</Service>
<Service Name="CallbackService" UID="{A5442363-EFFC-4B66-8AF4-5DA3231EA043}">
<Interfaces>
<Interface Name="Default" UID="{9EDDBC83-9C22-4EBC-B2EA-CA71DAC6A1AC}">
<Operations>
<Operation Name="GetGroupList" UID="{44DC024A-4A48-472A-89FA-25E8D88A500E}">
<Parameters>
<Parameter Name="Result" DataType="Binary" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="GetUserList" UID="{DF37CBE2-66B2-49FE-80D2-4BE46565F8FC}">
<Parameters>
<Parameter Name="Result" DataType="Binary" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="SendMsg" UID="{E704049E-213E-4D20-ACA2-BA0B33B5A336}">
<Parameters>
<Parameter Name="Result" DataType="Boolean" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
<Parameter Name="ARecver" DataType="Binary" Flag="In" >
</Parameter>
<Parameter Name="AMsg" DataType="Binary" Flag="In" >
</Parameter>
<Parameter Name="AMsgInfo" DataType="TMsgInfo" Flag="In" >
</Parameter>
<Parameter Name="AErrMsg" DataType="String" Flag="InOut" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="GetMsgList" UID="{512F646F-DD51-401C-843B-DFA03A1E245F}">
<Parameters>
<Parameter Name="Result" DataType="Binary" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="GetMsg" UID="{1A1DBF09-D98B-4735-9237-8D084105964C}">
<Parameters>
<Parameter Name="Result" DataType="Binary" Flag="Result">
</Parameter>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
<Parameter Name="AMsgInfo" DataType="TMsgInfo" Flag="In" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="GetOraSysDate" UID="{3ED6DF03-7344-46DC-8957-F0D0E0F29D89}">
<Parameters>
<Parameter Name="Result" DataType="String" Flag="Result">
</Parameter>
<Parameter Name="ACompCode" DataType="String" Flag="In" >
</Parameter>
</Parameters>
</Operation>
<Operation Name="SetMsgRead" UID="{67EC63B3-2F5F-4A1E-8FFE-E74D482DCADB}">
<Parameters>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In" >
</Parameter>
<Parameter Name="AMsgInfo" DataType="TMsgInfo" Flag="In" >
</Parameter>
</Parameters>
</Operation>
</Operations>
</Interface>
</Interfaces>
</Service>
</Services>
<EventSinks>
<EventSink Name="SrvCallbackEvent" UID="{643BEB92-C923-4C7E-9B3C-ABCEE86C6E34}">
<Interfaces>
<Interface Name="Default" UID="{50D06F25-6FBE-4924-9ABA-205DE720DD83}">
<Operations>
<Operation Name="UsersChange" UID="{118A707B-4F50-4863-AEC1-4E9A63F827F0}">
<Parameters>
<Parameter Name="AInfo" DataType="TLoginInfo" Flag="In">
</Parameter>
</Parameters>
</Operation>
<Operation Name="ShutdownServer" UID="{2E795EA2-C934-4626-943B-D37495B062E6}">
<Parameters>
<Parameter Name="AMessage" DataType="String" Flag="In">
</Parameter>
</Parameters>
</Operation>
<Operation Name="MsgChange" UID="{DC7B6234-E3A9-4A1F-BEBD-0AB35F6F835F}">
<Parameters>
<Parameter Name="AMsgInfo" DataType="TMsgInfo" Flag="In">
</Parameter>
</Parameters>
</Operation>
</Operations>
</Interface>
</Interfaces>
</EventSink>
</EventSinks>
<Structs>
<Struct Name="TLoginInfo" UID="{D932148B-9606-4615-97D5-793358785014}" AutoCreateParams="1">
<Elements>
<Element Name="SessionId" DataType="String">
</Element>
<Element Name="CompCode" DataType="String">
</Element>
<Element Name="CompName" DataType="String">
</Element>
<Element Name="UserId" DataType="String">
</Element>
<Element Name="UserName" DataType="String">
</Element>
<Element Name="Password" DataType="String">
</Element>
<Element Name="WorkClass" DataType="String">
</Element>
<Element Name="GroupName" DataType="String">
</Element>
<Element Name="HostName" DataType="String">
</Element>
<Element Name="IP" DataType="String">
</Element>
<Element Name="TermSId" DataType="String">
</Element>
<Element Name="TermSIP" DataType="String">
</Element>
<Element Name="TermSName" DataType="String">
</Element>
<Element Name="TermSPC" DataType="String">
</Element>
<Element Name="TermState" DataType="String">
</Element>
<Element Name="CompStr" DataType="String">
</Element>
<Element Name="Status" DataType="String">
</Element>
</Elements>
</Struct>
<Struct Name="TMsgInfo" UID="{A4625ADB-2D3B-4B95-BFC4-F91B8C1AFEF5}" AutoCreateParams="1">
<Elements>
<Element Name="CompCode" DataType="String">
</Element>
<Element Name="CompName" DataType="String">
</Element>
<Element Name="MsgPriority" DataType="String">
</Element>
<Element Name="MsgId" DataType="String">
</Element>
<Element Name="MsgSubject" DataType="String">
</Element>
<Element Name="MsgTime" DataType="String">
</Element>
<Element Name="MsgReply" DataType="String">
</Element>
<Element Name="MsgSenderId" DataType="String">
</Element>
<Element Name="MsgSenderName" DataType="String">
</Element>
<Element Name="MsgSenderWorkClass" DataType="String">
</Element>
<Element Name="MsgSenderWorkName" DataType="String">
</Element>
<Element Name="IsRead" DataType="Boolean">
</Element>
</Elements>
</Struct>
</Structs>
<Enums>
</Enums>
<Arrays>
</Arrays>
</Library>
