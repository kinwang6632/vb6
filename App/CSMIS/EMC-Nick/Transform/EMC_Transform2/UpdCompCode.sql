Prompt ***** 92.01.23 �ק�����N�XTable�����q�N�X    *****

Prompt  �����q�O
variable v_CompCode number;
begin
    select CodeNo into :v_CompCode from CD039 where rownum<=1;
end;

prompt  �ק�CD061 STB���M��]�N�X�����q�O
Update CD061 set CompCode=:v_CompCode;

prompt  �ק�CD060 �W�D���M��]�N�X�����q�O
Update CD060 set CompCode=:v_CompCode;

prompt  �ק�CD028 STB�t��N�X�ɪ����q�O
Update CD028 set CompCode=:v_CompCode;

COMMIT;

Prompt OK
exit




