

##################################################################
#
$       MODULE          = "dnasp_field";
$       AUTHOR          = "P. Willemin";
$       COPYRIGHT       = "NAGRAVision, (c) 1994-1998";
$       DATE            = "29-DEC-1998";
$       VERSION         = "v1.0";
#
##################################################################

@INCLUDES = ('<rw/rwtime.h>',
             '<rw/rwdate.h>');


$MAIN = 'FIELDS';

#------------------------------------------------------------------
#
#  FIELDS DEFINITION
#
#------------------------------------------------------------------


$EMM_D_T_FIELD          = "RWTime";
$EMM_DATE_FIELD         = "RWDate";

@FIELDS = (

        'EMM_D_T_FIELD',
        'EMM_DATE_FIELD'
);
