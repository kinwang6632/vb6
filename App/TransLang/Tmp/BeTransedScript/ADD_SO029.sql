/*
  �إ��v���[�c�� SO029
  Date: 2000.07.21
*/
set heading off
variable str varchar2(80)
exec :str := '<< �إ��v���[�c�� SO029 >> ' || to_char(sysdate, 'YYYY/MM/DD HH24:MI:SS');
spool p:\gird\Log\Add_SO029.log
print str

   DELETE FROM SO029;

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO0000', null, '�a�Q���u�q��CMIS', '', '' , null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1000', null, '�ȪA�޲z', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1100', null, '�Ȥ���/�ȪA�޲z', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11001', null, '��Ʒs�W', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11002', null, '��ƭק�', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO110021', null, '�ק�򥻸��', '', 'SO11002', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO110022', null, '�ק�b����', '', 'SO11002', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO110023', null, '�ק��L���', '', 'SO11002', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11003', null, '��Ƶ��P', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11005', null, '��ƦC�L', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11006', null, '�_�u�]�w', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1111', null, '��/�[/�_���A��', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11111', null, '��Ʒs�W', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11112', null, '��ƭק�', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11113', null, '��ƧR��', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11114', null, '��Ƭd��', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11115', null, '��ƦC�L', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11116', null, '���u����', '', 'SO1111', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1112', null, '���תA��', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11121', null, '��Ʒs�W', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11122', null, '��ƭק�', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11123', null, '��ƧR��', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11124', null, '��Ƭd��', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11125', null, '��ƦC�L', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11126', null, '���u����', '', 'SO1112', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1113', null, '��/��/�����A��', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11131', null, '��Ʒs�W', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11132', null, '��ƭק�', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11133', null, '��ƧR��', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11134', null, '��Ƭd��', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11135', null, '��ƦC�L', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11136', null, '���u����', '', 'SO1113', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1120', null, '�]�ƶ���', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11201', null, '��Ʒs�W', '', 'SO1120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11202', null, '��ƭק�', '', 'SO1120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11203', null, '��ƧR��', '', 'SO1120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11204', null, '��Ƭd��', '', 'SO1120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1131', null, '���O���س]�w', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11311', null, '��Ʒs�W', '', 'SO1131', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11312', null, '��ƭק�', '', 'SO1131', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11313', null, '��ƧR��', '', 'SO1131', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11314', null, '��Ƭd��', '', 'SO1131', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1132', null, '���O����s��', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11321', null, '��Ʒs�W', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11322', null, '��ƭק�', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11323', null, '��Ƨ@�o', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11324', null, '��Ƭd��', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11325', null, '��ƦC�L', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11326', null, '���u��ɳ�', '', 'SO1132', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1133', null, '���v���O����s��', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1141', null, '�ѽX���]�w', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1142', null, '�I�O�W�D�]�w', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11421', null, '��Ʒs�W', '', 'SO1133', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11422', null, '��ƭק�', '', 'SO1133', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11423', null, '��ƧR��', '', 'SO1133', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1150', null, 'PPV', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1161', null, '�����O���s��', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1162', null, '�Ȥ�A�ȥӧi', '', 'SO1100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11621', null, '��Ʒs�W', '', 'SO1162', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO11622', null, '��ƭק�', '', 'SO1162', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1200', null, '�a�}��ƺ޲z', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO12001', null, '��Ʒs�W', '', 'SO1200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO12002', null, '��ƭק�', '', 'SO1200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO12004', null, '��Ƭd��', '', 'SO1200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO12006', null, '���v���', '', 'SO1200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300', null, '�j�Ӹ�ƺ޲z', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13001', null, '��Ʒs�W', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13002', null, '��ƭק�', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13004', null, '��Ƭd��', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13005', null, '��ƦC�L', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13006', null, '���s�p��j�Ӥ�', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13007', null, '�Ȥ����O�վ�', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13008', null, '�w�]�X�����e�]�w', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO13009', null, '�s���Ȥ���', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300A', null, 'ñ��/���', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300B', null, '�e���X��', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300C', null, '�Φ���Ӧ�', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300D', null, '�Ӧ���Φ�', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300E', null, '�����e���޲z', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300E1', null, '��Ʒs�W', '', 'SO1300E', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300E2', null, '��ƭק�', '', 'SO1300E', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300E3', null, '��ƧR��', '', 'SO1300E', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300F', null, '�j�ӶO�v�]�w', '', 'SO1300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300F1', null, '��Ʒs�W', '', 'SO1300F', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300F2', null, '��ƭק�', '', 'SO1300F', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1300F3', null, '��ƧR��', '', 'SO1300F', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1400', null, '�P���I��ƺ޲z', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO14001', null, '��Ʒs�W', '', 'SO1400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO14002', null, '��ƭק�', '', 'SO1400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO14004', null, '��Ƭd��', '', 'SO1400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO14005', null, '��ƦC�L', '', 'SO1400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1500', null, '�ѹq���ƺ޲z', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15001', null, '��Ʒs�W', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15002', null, '��ƭק�', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15003', null, '��ƧR��', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15004', null, '��Ƭd��', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15005', null, '��ƦC�L', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15006', null, '�p�⤽��', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO150061', null, '��Ʒs�W', '', 'SO15006', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO150062', null, '��ƭק�', '', 'SO15006', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO150065', null, '��ƦC�L', '', 'SO15006', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO15007', null, '�C�L�I�ڳ�', '', 'SO1500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1600', null, '�w�}�t�κ޲z', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1610', null, '�ѽX���޲z', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO16101', null, '��Ʒs�W', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO16102', null, '��ƭק�', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO16103', null, '��ƧR��', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO16104', null, '��Ƭd��', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1620', null, '�ѽX���妸�]�w', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1630', null, 'IPPV�ϥ��I��Ū��', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1640', null, 'PPV�`�ت�', '', 'SO1610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO1700', null, '�A�ȥӧi�޲z', '', 'SO1000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO17002', null, '��ƭק�', '', 'SO1700', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO17004', null, '��Ƭd��', '', 'SO1700', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2000', null, '�u�Ⱥ޲z', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2100', null, '���u��n��', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2200', null, '���u��d��/�s��', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO22001', null, '��Ʒs�W', '', 'SO2200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO22002', null, '��ƭק�', '', 'SO2200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO22003', null, '��ƧR��', '', 'SO2200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO22004', null, '��Ƭd��', '', 'SO2200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO22005', null, '��ƦC�L', '', 'SO2200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2300', null, '���u��/����C�L', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2310', null, '���u��C�L', '', 'SO2300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2320', null, '���u����C�L', '', 'SO2300', 'SO2310B.RPT'); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2330', null, '���u������ӳ���C�L', '', 'SO2300', 'SO2310C.RPT'); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2400', null, '���u��鵲', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2500', null, '�G�ٰϰ�]�w', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO25001', null, '��Ʒs�W', '', 'SO2500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO25002', null, '��ƭק�', '', 'SO2500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO25003', null, '��ƧR��', '', 'SO2500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO25004', null, '��Ƭd��', '', 'SO2500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO25006', null, '���׾��^��', '', 'SO2500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2600', null, '�_�u�״_�޲z', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO2700', null, '�H����浲�ݦ���O��', '', 'SO2000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3000', null, '�p�O�޲z', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3100', null, '�X��޲z', '', 'SO3000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3110', null, '�������O��Ʋ���', '', 'SO3100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3120', null, '����������Ƭd�ߦC�L', '', 'SO3100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3121', null, '������������', '', 'SO3120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3122', null, '���������J�`', '', 'SO3120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3123', null, '���������ϰ�ιD�����R', '', 'SO3120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3124', null, '�����������ƤΪ��B���R', '', 'SO3120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3125', null, '�����X��ƤΪ��B�d��', '', 'SO3120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3130', null, '����վ�', '', 'SO3100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3200', null, '���O��޲z', '', 'SO3000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3210', null, '���O��L�J', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3220', null, '�L��Ǹ����s', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3230', null, '���O����d��', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3240', null, '���O��Ƭd��/�s��', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32401', null, '��Ʒs�W', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32402', null, '��ƭק�', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32403', null, '��Ƨ@�o', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32404', null, '��Ƭd��', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32405', null, '��ƦC�L', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO32406', null, '���u��ɳ�', '', 'SO3240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3250', null, '���վ�', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3251', null, '�L����B�վ�', '', 'SO3250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3252', null, '���O����@�o', '', 'SO3250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3253', null, '���O�洫��վ�', '', 'SO3250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3260', null, '��ڦC�L', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3261', null, '��妬�O��ڦC�L', '', 'SO3260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3262', null, '�浧���O��ڦC�L', '', 'SO3260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3263', null, '���O��ڭ��s�C�L', '', 'SO3260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3270', null, '�঩�b�Ϥ���Ʋ���', '', 'SO3200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3271', null, '��b���(�ۭq�榡)����', '', 'SO3270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3272', null, '�H�Υd���b��Ʋ���', '', 'SO3270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3273', null, '�{�d�N����Ʋ���', '', 'SO3270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3274', null, '����榡��b��Ʋ���', '', 'SO3270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3275', null, 'ATM�J�b������Ʋ���', '', 'SO3270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3300', null, '���ں޲z', '', 'SO3000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3310', null, '���O��Ƶn��', '', 'SO3300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3311', null, '���O��/������n��', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3312', null, '��b���\�Ϥ�(�ۭq�榡)�n��', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3313', null, '�H�Υd�J�b�n��', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3314', null, '�{�d�N���J�b�n��', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3315', null, '����榡�J�b�n��', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3316', null, 'ATM�J�b�n��', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO331A', null, '���u��Ҫ����O��Ƶn��', '', 'SO3310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3320', null, '���ڳ���d��', '', 'SO3300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3400', null, '�����ں޲z', '', 'SO3000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3410', null, '���O�楼����]�n��', '', 'SO3400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO3420', null, '�������@�~', '', 'SO3400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4000', null, '�b�Ⱥ޲z', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4100', null, '�L�b�B�z', '', 'SO4000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4110', null, '�鵲�b�@�~1(�t����)', '', 'SO4100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4120', null, '�鵲�b�@�~2(�t����)', '', 'SO4100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4130', null, '�o������(�t����)', '', 'SO4100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4140', null, '�ֳt�߱b�@�~', '', 'SO4100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4150', null, 'PPV�p�O�뵲(�t����)', '', 'SO4100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4200', null, '���B�z', '', 'SO4000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4210', null, '���C�a�b(�t����)', '', 'SO4200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4220', null, '�h���L�b�w����Ʀܾ��v��', '', 'SO4200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4230', null, '�R�������g���ʦ��O���ظ��', '', 'SO4200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO4240', null, '�g���ʦ��O���ت��B�վ�', '', 'SO4200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5000', null, '����޲z', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5100', null, '�H���Z�Ĳέp', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5200', null, '���O��Ʋέp', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5300', null, '�����W�D�έp', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5400', null, '�Ȥ�]�Ʋέp', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5500', null, '�Ȥ��Ʋέp', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5600', null, '���u��έp', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5700', null, '�U�ح�]�έp', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5800', null, '�Ȥ�A�Ȳέp', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5900', null, '�l�����Ҳέp', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5A00', null, '���U���R����', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5B00', null, '����ˮֳ���', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO5V00', null, 'PPV�έp', '', 'SO5000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6000', null, '�N�X�޲z', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6100', null, '�ϲեN�X��', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6110', null, '��F�ϥN�X', '', 'SO6100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61101', null, '��Ʒs�W', '', 'SO6110', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61102', null, '��ƭק�', '', 'SO6110', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61105', null, '��ƦC�L', '', 'SO6110', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6120', null, '�A�ȰϥN�X', '', 'SO6100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61201', null, '��Ʒs�W', '', 'SO6120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61202', null, '��ƭק�', '', 'SO6120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61205', null, '��ƦC�L', '', 'SO6120', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6130', null, '�u�{�եN�X', '', 'SO6100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61301', null, '��Ʒs�W', '', 'SO6130', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61302', null, '��ƭק�', '', 'SO6130', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61305', null, '��ƦC�L', '', 'SO6130', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6140', null, '���O�ϥN�X', '', 'SO6100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61401', null, '��Ʒs�W', '', 'SO6140', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61402', null, '��ƭק�', '', 'SO6140', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO61405', null, '��ƦC�L', '', 'SO6140', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6200', null, '���O�N�X��', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6210', null, '�Ȥ����O�N�X', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62101', null, '��Ʒs�W', '', 'SO6210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62102', null, '��ƭק�', '', 'SO6210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62105', null, '��ƦC�L', '', 'SO6210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6220', null, '�˾����O�N�X', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62201', null, '��Ʒs�W', '', 'SO6220', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62202', null, '��ƭק�', '', 'SO6220', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62205', null, '��ƦC�L', '', 'SO6220', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6230', null, '���ץӧi���O�N�X', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62301', null, '��Ʒs�W', '', 'SO6230', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62302', null, '��ƭק�', '', 'SO6230', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62305', null, '��ƦC�L', '', 'SO6230', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6240', null, '��������O�N�X', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62401', null, '��Ʒs�W', '', 'SO6240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62402', null, '��ƭק�', '', 'SO6240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62405', null, '��ƦC�L', '', 'SO6240', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6250', null, '�q�ܥӧi���O�N�X', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62501', null, '��Ʒs�W', '', 'SO6250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62502', null, '��ƭק�', '', 'SO6250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62505', null, '��ƦC�L', '', 'SO6250', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6260', null, '�C�餶�����O�N�X', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62601', null, '��Ʒs�W', '', 'SO6260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62602', null, '��ƭק�', '', 'SO6260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62605', null, '��ƦC�L', '', 'SO6260', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6270', null, '�P���I��~���O�N�X', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62701', null, '��Ʒs�W', '', 'SO6270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62702', null, '��ƭק�', '', 'SO6270', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62705', null, '��ƦC�L', '', 'SO6270', null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6280', null, '�P�P�N�X', '', 'SO6200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62801', null, '��Ʒs�W', '', 'SO6280', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62802', null, '��ƭק�', '', 'SO6280', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO62805', null, '��ƦC�L', '', 'SO6280', null); 

   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6300', null, '��]�N�X��', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6310', null, '�G�٭�]�N�X', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63101', null, '��Ʒs�W', '', 'SO6310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63102', null, '��ƭק�', '', 'SO6310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63105', null, '��ƦC�L', '', 'SO6310', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6320', null, '���P��]�N�X', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63201', null, '��Ʒs�W', '', 'SO6320', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63202', null, '��ƭק�', '', 'SO6320', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63205', null, '��ƦC�L', '', 'SO6320', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6330', null, '��ú�O��]�N�X', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63301', null, '��Ʒs�W', '', 'SO6330', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63302', null, '��ƭק�', '', 'SO6330', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63305', null, '��ƦC�L', '', 'SO6330', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6340', null, '�������]�N�X', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63401', null, '��Ʒs�W', '', 'SO6340', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63402', null, '��ƭק�', '', 'SO6340', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63405', null, '��ƦC�L', '', 'SO6340', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6350', null, '�h���]�N�X', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63501', null, '��Ʒs�W', '', 'SO6350', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63502', null, '��ƭק�', '', 'SO6350', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63505', null, '��ƦC�L', '', 'SO6350', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6360', null, '�u����]�N�X', '', 'SO6300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63601', null, '��Ʒs�W', '', 'SO6360', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63602', null, '��ƭק�', '', 'SO6360', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO63605', null, '��ƦC�L', '', 'SO6360', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6400', null, '�s���N�X��', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6410', null, '��D�s���N�X', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64101', null, '��Ʒs�W', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64102', null, '��ƭק�', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64104', null, '��Ƭd��', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64105', null, '��ƦC�L', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64106', null, '��D����', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64107', null, '��D���', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64108', null, '��D�έp', '', 'SO6410', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6420', null, '�Ȧ�s���N�X', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64201', null, '��Ʒs�W', '', 'SO6420', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64202', null, '��ƭק�', '', 'SO6420', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64205', null, '��ƦC�L', '', 'SO6420', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6430', null, '���O���إN�X', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64301', null, '��Ʒs�W', '', 'SO6430', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64302', null, '��ƭק�', '', 'SO6430', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64305', null, '��ƦC�L', '', 'SO6430', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6440', null, '�I�O�N�@�N�X', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64401', null, '��Ʒs�W', '', 'SO6440', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64402', null, '��ƭק�', '', 'SO6440', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64405', null, '��ƦC�L', '', 'SO6440', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6450', null, '�ؿv���A�N�X', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64501', null, '��Ʒs�W', '', 'SO6450', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64502', null, '��ƭק�', '', 'SO6450', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64505', null, '��ƦC�L', '', 'SO6450', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6460', null, '�~�W�s���N�X', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64601', null, '��Ʒs�W', '', 'SO6460', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64602', null, '��ƭק�', '', 'SO6460', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64605', null, '��ƦC�L', '', 'SO6460', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6470', null, '�ѧ˽s���N�X', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64701', null, '��Ʒs�W', '', 'SO6470', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64702', null, '��ƭק�', '', 'SO6470', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64705', null, '��ƦC�L', '', 'SO6470', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6480', null, '�W�D�s���N�X', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64801', null, '��Ʒs�W', '', 'SO6480', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64802', null, '��ƭק�', '', 'SO6480', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64805', null, '��ƦC�L', '', 'SO6480', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6490', null, '�Ȥ�ӷ��N�X', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64901', null, '��Ʒs�W', '', 'SO6490', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64902', null, '��ƭק�', '', 'SO6490', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64905', null, '��ƦC�L', '', 'SO6490', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64A0', null, '�A�Ⱥ��N�ץN�X', '', 'SO6400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64A01', null, '��Ʒs�W', '', 'SO64A0', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64A02', null, '��ƭק�', '', 'SO64A0', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO64A05', null, '��ƦC�L', '', 'SO64A0', null); 
/*
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6500', null, '��L������', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6510', null, '���O�覡�N�X', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6520', null, '�I�ں����N�X', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6530', null, '�|�v�N�X', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6540', null, '�R��覡�N�X', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6550', null, '�Ȥ᪬�A�N�X', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6560', null, '���u���A�N�X', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6570', null, '�H�Υd�����N�X', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6580', null, '���q�O�N�X', '', 'SO6500', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6590', null, '��Ʈw�N�X', '', 'SO6500', null); 
*/
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6600', null, '�@�Τp�ɺ��@', '', 'SO6000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6610', null, '�|�p��ؤp��', '', 'SO6600', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66101', null, '��Ʒs�W', '', 'SO6610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66102', null, '��ƭק�', '', 'SO6610', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6620', null, '���ƽs���p��', '', 'SO6600', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66201', null, '��Ʒs�W', '', 'SO6620', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66202', null, '��ƭק�', '', 'SO6620', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO6630', null, '�H�Ƹ�Ƥp��', '', 'SO6600', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66301', null, '��Ʒs�W', '', 'SO6630', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO66302', null, '��ƭק�', '', 'SO6630', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7000', null, '�t�κ޲z', '', 'SO0000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7200', null, '�ϥΪ̺޲z', '', 'SO7000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7210', null, '�ϥΪ̷s�W�R��', '', 'SO7200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO72101', null, '��Ʒs�W', '', 'SO7210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO72102', null, '��ƭק�', '', 'SO7210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO72103', null, '��ƧR��', '', 'SO7210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO72105', null, '��ƦC�L', '', 'SO7210', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7220', null, '�ϥΪ��v���]�w', '', 'SO7200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7230', null, '�ϥΪ��v���ƻs', '', 'SO7200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7240', null, '�ϥΪ��v������C�L', '', 'SO7200', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7300', null, '�t�ΰѼƳ]�w', '', 'SO7000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7310', null, '�~�̰ѼƳ]�w', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7320', null, '���u�ѼƳ]�w', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7330', null, '�w�]���O�ѼƳ]�w', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7340', null, '��L�w�]�ѼƳ]�w', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7350', null, '��X�ѼƳ]�w', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7360', null, '�w�����u�ɬq�]�w', '', 'SO7300', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7400', null, '�a�}������ƺ޲z', '', 'SO7000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74001', null, '��Ʒs�W', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74002', null, '��ƭק�', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74003', null, '��ƧR��', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74004', null, '��Ƭd��', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74005', null, '��ƦC�L', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO74006', null, '���]�w', '', 'SO7400', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7500', null, '�妸�@�~', '', 'SO7000', null); 
/*
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO7600', null, '����榡�ɳ]�w', '', 'SO7000', null); 
*/
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8000', null, '��L', '', 'SO0000', null);
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8100', null, '���G��', '', 'SO8000', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO81001', null, '��Ʒs�W', '', 'SO8100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO81002', null, '��ƭק�', '', 'SO8100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO81003', null, '��ƧR��', '', 'SO8100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO81005', null, '��ƦC�L', '', 'SO8100', null); 
   INSERT INTO SO029 (Mid, Layer, Prompt, Command, Link, FormatName)
              VALUES ( 'SO8200', null, '������Ʈw', '', 'SO8000', null); 
@\gird\script\add_so029_1;
   update so029 set Group1=1;

-- ************************
-- �s�W�@���ťո�Ʃ�SO039 ���O��Ƶ���
insert into SO039 (Memo1, Memo2) values (null, null);

COMMIT;

spool off
set heading on
