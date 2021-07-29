options (errors=200)
load data
infile 'CD024_Map.txt'
append into table CD024
fields terminated by X'09' TRAILING NULLCOLS
(CodeNo CHAR "substrb(rtrim(:CodeNo),1,3)"
,Description CHAR "substrb(rtrim(:Description),1,20)"
,ChannelID CHAR "substrb(rtrim(:ChannelID),1,12)"
,CompCode DECIMAL EXTERNAL
,StopFlag DECIMAL EXTERNAL
,ChanceDays DECIMAL EXTERNAL
)
