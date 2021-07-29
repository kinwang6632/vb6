

#--+-------------------------------------------------------------------------+--
#--+ package shared_data
#--+-------------------------------------------------------------------------+--
#--+ Author              : Patrice Willemin                                  +--
#--+ CopyRight           : NAGRAVision, (c) 1994-1998                        +--
#--+ Creation            : 15-JUN-1998                                       +--
#--+ Update              :                                                   +--
#--+ Version             : 1.0                                               +--
#--+ Previous version    : /                                                 +--
#--+                                                                         +--
#--+ Description                                                             +--
#--+                     : Provide routine to handle shared perl variable
#--+-------------------------------------------------------------------------+--

package shared_data;


# Global variable area

$SEMID = 0 ;			# Semaphore identifier
$SEMNR = 0 ;			# Semaphore number
$IPC1  = "IPC1_SHD" ;		# IPC channel 1
$IPC2  = "IPC2_SHD" ;		# IPC channel 2
$is_started = 0;		# Controller not started initially

require FTcp;
require Ipc;
require Trace;

sub lock {
  local ($sts) = Ipc::sem_wait ($SEMID, $SEMNR) ;
  Trace::put ("sem_wait fails; $!;$sts") if $!;
}

sub unlock {
  local ($sts) = Ipc::sem_post ($SEMID, $SEMNR) ;
  Trace::put ("sem_post fails; $!;$sts") if $!;
}
  

sub get {

  # Description
  #   Get a parameter.
  #
  # Synopsis
  #   get (name) return value.

  local ($name) = @_ ;
  local ($data) ;

  lock ;
  FTcp::send_data ($IPC1,"G $$-$tid $name");
  $data = FTcp::receive_data ($IPC1);
  unlock ;

  $data =~ s/^(\S*)\s//;
  Trace::put ("mismatch transid $1,$$-$tid") if ($1 ne "$$-$tid");
  $tid++;

  return $data ;
}

sub set {

  # Description
  #   Set a parameter.
  #
  # Synopsis
  #   set (name, value)

  local ($name, $value) = @_ ;
  local ($data) ;
 
  lock ;
  FTcp::send_data ($IPC1,"S $$-$tid $name\1$value");  
  unlock ;

  $tid++;
  return 1;
}

sub del {
  
  # Description
  #   Delete a variable.
  #
  # Synopsis 
  #   del (name)

  local ($name) = @_ ;
  local ($data) ;

  lock ;
  FTcp::send_data ($IPC1, "D $$-$tid $name") ;
  unlock ;

  $tid++;
  return 1;
}


sub select_table {

  # Description
  #   Get the entire associative array
  #
  # Synopsis
  #   select_table

  local ($selection) = @_ ;
  local ($value,$data,$key);
  local (%params);

  # First check the selection to prevent the abortion of the suprocess 
  # while in the critical section, in case of a fatal pattern matching syntax
  # error.

  /$selection/ ;

  lock;
  FTcp::send_data ($IPC1, "T $$-tid $selection") ;
  $tid++;
  while ($data = FTcp::receive_data ($IPC1)) {
    ($key, $value) = split ("\1",$data);
    $params {$key} = $value ;
  }
  unlock;

  return %params;
}


sub start_controller {

  # Description
  #   Handles the parameter operation commands (Modification, Increment,
  #   Deletion and Get operations)
  #   This routine is performed by a specific subprocess.
  #
  # Synopsis
  #   start_controller ([data file], [sem_number=10])

  ($SEMID, $SEMNR) = @_ ;

  die 'shared_data.start_controller: invalid argument number' 
      if !defined $SEMID || !defined $SEMNR ;

  local ($data, $value, $key, %ARR, $child);

  FTcp::create_link ($IPC1, $IPC2) ;

  if (!($child = fork())) {		# Start the ARRAY manager task

    close $IPC1;
    
    RECEIVER: 

    while (1) {
      $data = FTcp::receive_data ($IPC2) ;

      #  Remove comment for DEBUG  
      ## print "IPC2 receive: $data\n";

      $data =~ s/^(\w) (\S*)\s//;
      $_ = $1 ;
      $id = $2 ;

      CASE: {
        /G/ && FTcp::send_data ($IPC2, $id . " " . $ARR {$data});
        /S/ && (($key,$value) = split ("\1",$data), $ARR {$key} = $value);  
        /D/ && delete $ARR {$data};
        /T/ && do {
              $data = '.' if $data eq '';
              foreach $key (keys (%ARR)) {
                FTcp::send_data ($IPC2, "$key\1$ARR{$key}") if $key =~ /$data/;
              }
              FTcp::send_data ($IPC2, "") ;
            };
        /X/ && last RECEIVER;
      }
    }
    close $IPC2 ;
    exit ;
  }

  close $IPC2 ;
  $is_started  = 1;
  return $child;
}


sub stop_controller {
  
  lock ;
  FTcp::send_data ($IPC1, "X ") ;
  unlock ;
  close ($IPC1) ;
  $is_started = 0;
}

sub controller_started {
  return $is_started;
}

1;
