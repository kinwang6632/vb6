

##################################################################
#
$       MODULE          = "common_field";
$       AUTHOR          = "P. Willemin";
$       COPYRIGHT       = "NAGRAVision, (c) 1994-1998";
$       DATE            = "29-DEC-1998";
$       VERSION         = "v1.0";
#
##################################################################

push (@MODULE_LIST, $MODULE);


$MAIN = 'FIELDS';

#------------------------------------------------------------------
#
#  FIELDS DEFINITION
#
#------------------------------------------------------------------


$INT_FIELD          = "int";
$REV_INT_FIELD	    = "int";
$STRING_FIELD       = "C_VBytes";

@FIELDS = (

        'INT_FIELD',
        'REV_INT_FIELD',
        'STRING_FIELD'
);
