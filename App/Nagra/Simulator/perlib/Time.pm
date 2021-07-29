
#--+-------------------------------------------------------------------------+--
#--+ package Time
#--+-------------------------------------------------------------------------+--
#--+ Author              : Patrice Willemin                                  +--
#--+ CopyRight           : NAGRAVision, (c) 1994-1998                        +--
#--+ Creation            : 10-JUN-1998                                       +--
#--+ Update              :                                                   +--
#--+ Version             : 1.0                                               +--
#--+ Previous version    : /                                                 +--
#--+                                                                         +--
#--+ Description                                                             +--
#--+                     : Provide time related routines
#--+-------------------------------------------------------------------------+--

package Time ;


%MONTH = (
    '1','JAN',
    '2','FEB',
    '3','MAR',
    '4','APR',
    '5','MAY',
    '6','JUN',
    '7','JUL',
    '8','AUG',  
    '9','SEP',
    '10','OCT',
    '11','NOV',
    '12','DEC') ;

%MINDEX = (
    'JAN',1,
    'FEB',2,
    'MAR',3,
    'APR',4,
    'MAY',5,
    'JUN',6,
    'JUL',7,
    'AUG',8,
    'SEP',9,
    'OCT',10,
    'NOV',11,
    'DEC',12) ;
     
sub image {
  local ($time,$format,$gmt) = @_ ;
 
  $format = 'VMS' if !defined $format;

  local ($sec, $min, $hour, $mday, $mon, $year, $wday, $yday, $isdst) =
    $gmt eq 'GMT' ? gmtime ($time) : localtime ($time);

  $year = $year < 90 ? $year+2000 : $year+1900;
  $mon += 1;

  return sprintf "%2.2d-%s-%s %2.2d:%2.2d:%2.2d",$mday,$MONTH{$mon},
    $year, $hour,$min,$sec 
    if $format eq 'VMS';

  return sprintf "%4.4d%2.2d%2.2d%2.2d%2.2d%2.2d",
    $year,$mon,$mday,$hour,$min,$sec 
    if $format eq 'COMPACT';

 
  # handle customized format

  local ($YYYY,$MM,$DD,$hh,$mm,$ss) = split ('\.',
      sprintf ("%4.4d.%2.2d.%2.2d.%2.2d.%2.2d.%2.2d",
               $year,$mon,$mday,$hour,$min,$sec)) ;

  $format =~ s/YYYY/$YYYY/;
  $format =~ s/MMM/$MONTH{$mon}/;
  $format =~ s/MM/$MM/;
  $format =~ s/DD/$DD/;
  $format =~ s/hh/$hh/;
  $format =~ s/mm/$mm/;
  $format =~ s/ss/$ss/;

  if ($format =~ m/YY/) {
    $year = substr ($YYYY,2);
    $format =~ s/YY/$year/;
  }
  return $format;
}


sub delta_image {
  local ($time1, $time2) = @_ ;
  local ($dtime) = $time1 - $time2 ;
  local ($sign) = ' ' ;
  local ($day) = '';

  if ($time1 < $time2) { 
    $sign ='-' ;
    $dtime = -$dtime;
  }

  if (int($dtime / 86400) > 0) {
    $day = sprintf "%d-",$dtime / 86400;
  }

  sprintf "%s%s%2.2d:%2.2d:%2.2d",
       $sign,
       $day,
       int($dtime / 3600) % 24,
       int($dtime / 60) % 60,
       $dtime % 60 ;
}
1;

sub split_date {    
  local ($date, $format) = @_ ;
  $format = 'VMS' if !defined $format;
 
  if ($format eq 'VMS') { 
    $date =~ /(\d\d)-(\w\w\w)-(\d\d\d\d) (\d\d):(\d\d):(\d\d)/ ;
    return ($3,sprintf ("%2.2d",$MINDEX{$2}),$1,$4,$5,$6);
  }

  if ($format eq 'UNIX') {

    local ($sec, $min, $hour, $mday, $mon, $year) = localtime (time);
    $year = $year < 90 ? $year+2000 : $year+1900;

    $date =~ tr/a-z/A-Z/;
    $date =~ /(\w\w\w)\s+(\d+)\s+(\d\d):(\d\d)/ && 
      return ($year,sprintf ("%2.2d",$MINDEX{$1}),
                     sprintf ("%2.2d",$2),$3,$4,'00');

    $date =~ /(\w\w\w)\s+(\d+)\s+(\d\d\d\d)/ &&
      return ($3,sprintf ("%2.2d",$MINDEX{$1}),
                 sprintf ("%2.2d",$2),'00','00','00');
  }

  $format =~ s/YYYY/(\\d\\d\\d\\d)/;
  $format =~ s/YY/(\\d\\d)/;
  $format =~ s/MMM/(\\w\\w\\w)/;
  $format =~ s/MM/(\\d\\d)/;
  $format =~ s/DD/(\\d\\d)/;
  $format =~ s/hh/(\\d\\d)/;
  $format =~ s/mm/(\\d\\d)/;
  $format =~ s/ss/(\\d\\d)/;

  $date =~ /$format/;
  
  return ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10);

}
