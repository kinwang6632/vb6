Prompt ***** 92.01.23 修改相關代碼Table的公司代碼    *****

Prompt  取公司別
variable v_CompCode number;
begin
    select CodeNo into :v_CompCode from CD039 where rownum<=1;
end;

prompt  修改CD061 STB結清原因代碼的公司別
Update CD061 set CompCode=:v_CompCode;

prompt  修改CD060 頻道結清原因代碼的公司別
Update CD060 set CompCode=:v_CompCode;

prompt  修改CD028 STB配件代碼檔的公司別
Update CD028 set CompCode=:v_CompCode;

COMMIT;

Prompt OK
exit




