
#--+-------------------------------------------------------------------------+--
#--+ package Msg
#--+-------------------------------------------------------------------------+--
#--+ Author              : Patrice Willemin                                  +--
#--+ CopyRight           : NAGRAVision, (c) 1994-1998                        +--
#--+ Creation            : 03-JUN-1998                                       +--
#--+ Update              :                                                   +--
#--+ Version             : 1.0                                               +--
#--+ Previous version    : /                                                 +--
#--+                                                                         +--
#--+ Description                                                             +--
#--+                     : Provide routines to signal messages of which
#--+			   the definition is given by an ODL unit file.
#--+			   In this version, the following feature
#--+			   are supported:
#--+
#--+			   DEVICE kind
#--+			     NULL_DEVICE, DEFAULT_OUTPUT, TEXT, NETWORK,
#--+			     SHARED_TEXT, APPEND_SHARED_TEXT.
#--+		
#--+	                   DEVICE attributes
#--+			     DEVICE_ID, DEVICE_NAME, DEVICE_KIND,
#--+                         MAX_RECORD,
#--+                         MAX_VERSION, MIN_LEVEL, MIN_SEVERITY, FORMAT
#--+
#--+			   FACILITY attributes
#--+                         FACILITY_ID, DEVICES, FACILITY_NAME,
#--+                         MIN_LEVEL, FORMAT
#--+-------------------------------------------------------------------------+--

package Msg;

require File_mgr;
require 'ctime.pl';
require FTcp;
require Time;

$PM_SERVICE = 0 ;
$PM_NODE = '';

$gbl_access_mode = 'direct';
$gbl_table_name = '';
$gbl_connect_ok = 0;
$gbl_next_try = 0;

$NO_TRACE = 1 ;
$SEMID = 0 ;
$SEMNR = 0 ;
$IPC1 = 'IPC1_MSG' ;
$IPC2 = 'IPC2_MSG' ;


sub connect {
  ($PM_SERVICE, $PM_NODE) = @_ ;
  $PM_NODE = 'localhost' if ! defined $PM_NODE ;
  $gbl_access_mode = 'remote';
  $gbl_connect_ok = FTcp::connect_host (*PM_LINK, 'localhost', $PM_SERVICE);
  if (!$gbl_connect_ok) {
    printf "%s Connect fail to service $PM_SERVICE\n",Time::image(time);
  }
}

sub disconnect {
 FTcp::close (*PM_LINK) ;
 $gbl_access_mode = 'local';
} 


sub load_module {
 
  # Description
  #   Load the given ODL unit into the associative arrays
  #   DEVICE, FACILITY and MESSAGE.
  #
  # Synopsis
  #   load_module (module)


  local ($module) = @_;
  local ($done,$TBL);
  local ($tmp_file) = "$ENV{'tmp'}/oda_$$.tmp";


  # Due to a PERL pssibly bug, the instruction below (in comment) 
  # must be avoided. Some times. The end of file appears whereas not
  # all the data have been read. 
  #
  # open (ODA, "oda $module|") 
  
  system ("oda $module > $tmp_file") ;
  open (ODA, "$tmp_file") ;
  while (<ODA>) {

    $done=0;
    if (/\/(\w*)\/ (\w+)\$OBJ\$(\w+)\$(\w+)\$(\d+)\s*=\s*\/(.*)\/$/) {
      ($_,$key,$value) = ($2,"$1.$3.$4.$5", $6);
      $TBL=$1;
      $done=1
    }
    elsif (/\/(\w*)\/ (\w+)\$OBJ\$(\w+)\$(\w+)\s*=\s*\/(.*)\/$/) {
      ($_,$key,$value) = ($2,"$1.$3.$4", $5);
      $TBL=$1;
      $done=1
    }

    if ($done) {
      CASE: {
            /MESSAGE/  && do {$MESSAGE  {$key} = $value};
            /DEVICE/   && do {$DEVICE   {$key} = $value};
            /FACILITY/ && do {$FACILITY {$key} = $value};
      }
    }
  }
  close (ODA);
  unlink $tmp_file;
  $gbl_table_name = $TBL ;
  # print STDOUT "Module $module loaded\n";  # Remove comment for DEBUG
}


sub set_table {

  # Description
  #   Set the default table name

  ($gbl_table_name) = @_;
}


sub dump {

  # Description
  #   Dump the given object table to the default output. (To be used
  #   for debugging purpose.
  #
  # Synopsis
  #   dump (object)	; object = MESSAGE | FACILITY | DEVICE


  local ($object) = @_ ;
  local (%TABLE,$name) ;

  eval "\%TABLE = \%$object";
  foreach $name (sort keys (%TABLE)) {
    printf "%-40s%-20s\n", $name, $TABLE{$name};
  }
}


sub device {

  # Description
  #   Defined the device to which the message has to be sent.
  #   Must be used as the last argument of any argument message.

  return "\2@_";
}


sub severity_order {
  local ($severity) = @_ ;
  local (@SEV_TABLE) = ('SUCCESS','INFORMATION','WARNING','ERROR','FATAL_ERROR');
  local ($pos) = 1;

  foreach (@SEV_TABLE) {
    return $pos if /$severity/;
    $pos++
  }
  return 0; # For any undefined severity
}
  

sub form {
  local ($TBL,$dev,$MSG, $SEV, $FAC, $TXT) = @_ ;
  local ($M, $TIME, $FAC_name, $year, $month, $day, $time, $week_day);

  local (%MONTH) = (
    "JAN","01","FEB","02","MAR","03","APR","04","MAY","05","JUN","06",
    "JUL","07","AUG","08","SEP","09","OCT","10","NOV","11","DEC","12") ;

  $FAC_name = $FACILITY {"$TBL.$FAC.FACILITY_NAME"};
  ($week_day, $month, $day, $time, $year) = split (' ',ctime(time));
  $month =~ y/a-z/A-Z/;
  $TIM = sprintf "%2.2d-%s-%4.4d %s", $day,$month,$year,$time ;

  if (! ($M = $DEVICE {"$TBL.$dev.FORMAT"})) {
    $M = $FACILITY {"$TBL.$FAC.FORMAT"}
  }
 
 
  # Special TEXT symbols remplacement.

  $TXT =~ s/%p/$$/;	# Process identifier
  $TXT =~ s/%n/$0/;	# Program name
  $TXT =~ s/%u/$>/;	# User identifier
    
  $M =~ s/%FAC/$FAC_name/;
  $M =~ s/%SEV/$SEV/;
  $M =~ s/%COD/$MSG/; 
  $M =~ s/%TXT/$TXT/;
  $M =~ s/%TIM/$TIM/;
  $M =~ s/%%/%/;
  $M =~ m/%FTM<([^>]+)>/;

  if ($T = $1) {
    ($hour, $minute, $second) = split (":",$TIM);
    $day = sprintf "%2.2d",$day;
    
    $T =~ s/YYYY/$year/;
    $T =~ s/MMM/$month/;
    $T =~ s/MM/$MONTH{"$month"}/;
    $T =~ s/DD/$day/;
    $T =~ s/hh/$hour/;
    $T =~ s/mm/$minute/;
    $T =~ s/ss/$second/;
    if ($T =~ m/YY/) {
      $year %= 100;
      $T =~ s/YY/$year/;
    }
    $M =~ s/%FTM<[^>]+>/$T/;
  }

  return $M ; 
}


sub put_message_local {
  
  local ($TBL,$MSG,@params) = @_ ;

  Ipc::sem_wait ($SEMID, $SEMNR) ;
  FTcp::send_data ($IPC1, "$TBL.$MSG " . join ("\1",@params), $NO_TRACE);
  Ipc::sem_post ($SEMID, $SEMNR) ;

}


sub put_message_remote {

  local ($TBL,$MSG,@params) = @_ ;
  local ($msgcnt,$answer) ;

  if (!$gbl_connect_ok && time > $gbl_next_try) {
    $gbl_connect_ok = FTcp::connect_host (*PM_LINK, 'localhost', $PM_SERVICE);
    if ($gbl_connect_ok) {
      printf "%s Connect to service $PM_SERVICE\n",Time::image(time);
    }
    $gbl_next_try = time + 10;
  }

  if ($gbl_connect_ok) {
    $msgcnt = FTcp::send_data (*PM_LINK, "$TBL.$MSG " . join ("\1",@params));
    $answer = FTcp::receive_data (*PM_LINK) ;
    $gbl_connect_ok=0 if !defined $msgcnt || !defined $answer ; 
    return $answer if $gbl_connect_ok ;
  }

  if (!$gbl_connect_ok) {
    print STDOUT "%???-?-$MSG, ??? (" . join (',',@params) . ")\n";
    return 1; # Failure, message not sent
  }
  return 0; # Success
}


sub open_device {

  local ($TBL,$dev) = @_ ;
  local ($mode, $dir, $body, $ver, $DVN) ;
  local ($app) = $0;
    
  $_ = $DEVICE{"$TBL.$dev.DEVICE_KIND"} ;
  return 0 if ! $_ ;

  CASE: {

    /DEFAULT_OUTPUT/ && do {
      $DEV_DESC{"$TBL.$dev"} = STDOUT ;
      };

    /NETWORK/ && do {
      $DVN = $DEVICE{"$TBL.$dev.DEVICE_NAME"};
      $DVN =~ /([^\/]*)\/(\w+)\.(\w+)/;
      
      local ($node, $service, $protocol) = ($1, $2, $3);

      if ($protocol eq 'FTCP') {
        FTcp::connect_host ($dev, $node, $service) || print "unable to access $DVN\n";
        $DEV_DESC{"$TBL.$dev"} = $dev;  
      } 
      else {
        print "unable to access $DVN; protocol not supported\n";
      }
      };


    /TEXT/ && do {
      $mode = "> " ;

      $DVN = $DEVICE{"$TBL.$dev.DEVICE_NAME"};

      $app =~ s/.pl//;
      $DVN =~ s/%p/$$/;		# Process identifier
      $DVN =~ s/%n/$app/;	# Program name
      $DVN =~ s/%u/$>/;		# User identifier

      ($dir, $body) = File_mgr::components ($DVN);
      $ver = 1 + File_mgr::last_version ($dir, $body);

      /APPEND/ && do {$mode = ">> "; $ver--};

      # Remove comment for DEBUG
      ## print "open device $dev = $DVN-$ver\n";

      open ($dev, $mode . $DVN . "-$ver") || print "unable to open $DVN-$ver\n";

      /SHARED_TEXT/ && do {select ((select ($dev), $|=1)[$[])};
      $DEV_DESC{"$TBL.$dev"} = $dev;  
      $RECORD_COUNT{"$TBL.$dev"} = 0;
      File_mgr::purge ($dir,$body,$DEVICE{"$TBL.$dev.MAX_VERSION"})
        if $DEVICE{"$TBL.$dev.MAX_VERSION"} > 0 ;
      }
  }
  return 1; # Success
}  


sub put_message_direct {

  # Description
  #   Get the full description of the message and put it to the 
  #   corresponding list of target devices. A message consists of
  #   a mnemonic (the message_id), a facility, a severity, a level and
  #   a short message description (text)
  #
  # Synopsis
  #   put (message_id, {param,...})


  local ($TBL,$MSG,@params) = @_;
  local ($dev_index) = 1;
  local ($LEVEL,$SEV,$FAC,$file,$msgcnt,$answer,$maxrec,$dev,$dev_status);
  local ($single_dev) = 0;
  local ($status);

  if ($params[$#params] =~ /^\x02(.*)/) {
    $dev = $1;
    $#params -=1;
  }

  $single_dev = 1 if defined $dev;

  $SEV = substr ($MESSAGE {"$TBL.$MSG.SEVERITY"},0,1);


  if ($SEV eq "") {
    print STDOUT "%???-?-$MSG, ??? (" . join (',',@params) . ")\n";
    return 1; # Failure, message not defined
  }

  $FAC = $MESSAGE {"$TBL.$MSG.FACILITY"};
  $TXT = $MESSAGE {"$TBL.$MSG.TEXT"};

  if (! ($LEVEL = $MESSAGE{"$TBL.$MSG.LEVEL"})) {
    $LEVEL = $FACILITY {"$TBL.$FAC.DEFAULT_LEVEL"};
  }

  return 0 if $LEVEL < $FACILITY {"$TBL.$FAC.MIN_LEVEL"} ;  # Success

  foreach (@params) {
    $TXT =~ s/%/$_/
  }

  while (1) {
    $dev = $FACILITY {"$TBL.$FAC.DEVICES.${dev_index}"} if !$single_dev;

    last if !$dev;
    $dev_index++ ;
   
    $dev_status=1;
    $dev_status = open_device ($TBL, $dev) if !$DEV_DESC{"$TBL.$dev"};

    if (!$dev_status) {
      print STDOUT form ($TBL,'',$MSG,$SEV,$FAC,$TXT)," (???DEVICE $dev)\n";
    }
    
 
    if (severity_order ($MESSAGE{"$TBL.$MSG.SEVERITY"}) >= 
        severity_order ($DEVICE{"$TBL.$dev.MIN_SEVERITY"}) || 
        $LEVEL >= $DEVICE{"$TBL.$dev.MIN_LEVEL"}) {
  
      $file = $DEV_DESC{"$TBL.$dev"} ;

      $dev_kind = $DEVICE{"$TBL.$dev.DEVICE_KIND"} ;
       
      ## print "###send data to $file of kind $dev_kind\n"    ;

      if ($DEVICE{"$TBL.$dev.DEVICE_KIND"} eq 'NETWORK') {
        FTcp::send_data ($file, form ($TBL,$dev,$MSG,$SEV,$FAC,$TXT));
      }
      else {
        print $file form ($TBL,$dev,$MSG,$SEV,$FAC,$TXT),"\n";
        $RECORD_COUNT {"$TBL.$dev"}++;

        $maxrec = $DEVICE{"$TBL.$dev.MAX_RECORD"} ;

        if ($RECORD_COUNT{"$TBL.$dev"} > $maxrec && defined $maxrec) {
          close ($file);
          open_device ($TBL, $dev) ;
        }
      }
    
    }

    last if $single_dev;

  } 
  return 0; # Success
}


sub image {

  local ($MSG,@params) = @_;
  local ($dev_index) = 1;
  local ($LEVEL,$SEV,$FAC,$file,$ver,$mode,$dir,$body);

  if ($MSG =~ /(\w*)\.(\w+)/) {
    ($TBL, $MSG) = ($1,$2) ;
  }
  else {
    $TBL = $gbl_table_name ;
  }

  $SEV = substr ($MESSAGE {"$TBL.$MSG.SEVERITY"},0,1);
  $FAC = $MESSAGE {"$TBL.$MSG.FACILITY"};
  $TXT = $MESSAGE {"$TBL.$MSG.TEXT"};

  return "%???-?-$MSG, ??? (" . join (',',@params) . ")" if !$SEV ;

  if (! ($LEVEL = $MESSAGE{"$TBL.$MSG.LEVEL"})) {
    $LEVEL = $FACILITY {"$TBL.$FAC.DEFAULT_LEVEL"};
  }

  return "" if $LEVEL < $FACILITY {"$TBL.$FAC.MIN_LEVEL"} ;

  foreach (@params) {
    $TXT =~ s/%/$_/
  }
  return form ($TBL,"",$MSG,$SEV,$FAC,$TXT);
}

sub put {

  # Description
  #   Get the full description of the message and put it to the 
  #   corresponding list of target devices. A message consists of
  #   a mnemonic (the message_id), a facility, a severity, a level and
  #   a short message description (text)
  #
  # Synopsis
  #   put (message_id, {param,...}, [target device])

  local ($MSG, @params) = @_;
  local ($TBL) ;
  local ($sts) ;

  if ($MSG =~ /(\w*)\.(\w+)/) {
    ($TBL, $MSG) = ($1,$2) ;
  }
  else {
    $TBL = $gbl_table_name ;
  }

  $_ = $gbl_access_mode ;
  CASE : {
    /remote/ && {$sts = put_message_remote ($TBL, $MSG, @params)} ;
    /local/  && {$sts = put_message_local  ($TBL, $MSG, @params)} ;
    /direct/ && {$sts = put_message_direct ($TBL, $MSG, @params)} ;
  }

  return $sts ;
}



sub start_controller {

  ($SEMID, $SEMNR) = @_ ;

  die 'Msg.start_controller: invalid argument number'
      if !defined $SEMID || !defined $SEMNR ;

  local ($data, $child, @args, $argline,$dev) ;

  FTcp::create_link ($IPC1, $IPC2) ;
  if (!($child = fork())) {             # Start the controller task

    close $IPC1;

    RECEIVER:

    while (1) {
      $data = FTcp::receive_data ($IPC2, $NO_TRACE) ;
      if ($data =~ /(\w*)\.(\w+) (.*)/ || $data =~ /(\w*)\.(\w+)$/) {
        ($argline, $dev) = split ("\2",$3);
        @args = split ("\1",$argline) ;
        
        put_message_direct ($1,$2,@args,$dev) ;
      }
    }
    close $IPC2;
    exit;
  }
  close $IPC2;
  $gbl_access_mode = 'local';
  return $child;
}

1;
