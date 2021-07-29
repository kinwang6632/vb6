

#--+-------------------------------------------------------------------------+--
#--+ package Tcp
#--+-------------------------------------------------------------------------+--
#--+ Author              : Patrice Willemin                                  +--
#--+ CopyRight           : NAGRAVision, (c) 1994-1998                        +--
#--+ Creation            : 03-JUN-1998                                       +--
#--+ Update              :                                                   +--
#--+ Version             : 1.0                                               +--
#--+ Previous version    : /                                                 +--
#--+                                                                         +--
#--+ Description                                                             +--
#--+                     : Provide Tcp communication primitives
#--+-------------------------------------------------------------------------+--

  
package Tcp;

# Package summary
#
# -- create_service (Out socket, port number)
# -- delete_service (socket)
# -- accept_call (socket, Out subsocket)
# -- connect_host (Out socket, target address, port number)
# -- receive_data (socket, [max data size]) return data
# -- send_data (socket, data) return nb.bytes sent
# -- close_link (socket)

require Ipc;
require Time;
require Trace;
require Msg;
 
# TCP related constants
#
$AF_INET	= 2 ;
$AF_UNIX	= 1 ;
$SOCK_STREEM	= 1 ;
$sockaddr 	='S n a4 x8';
$SOL_SOCKET	= 0xFFFF ;
$SO_REUSEADDR	= 0x0004 ;
$LOCAL          = pack("cccc",127,0,0,1);

# Global constants
#
$BUF_SIZE	  = 16384 ;
$MAX_PENDING_CALL = 5 ;

$DEBUG = 0 ;

sub set_trace_mode {
  
  # Description
  #   Set the trace mode for TCP communication.
  #
  # Synopsis
  #   set_trace_mode (mode);

  local ($mode) = @_;
  $MODE = $mode;
  $DEBUG = 1;
}

sub trace {
  local ($text) = @_ ;
  local ($s,$digit) ;

  if ($DEBUG) {
    foreach $digit (unpack ('C*',$text)) {
      if ($digit < 32 || $digit > 127) {
        $s .= sprintf ("<%X>",$digit) ;
      }
      else {
        $s .= sprintf ("%c",$digit);
      } 
    }
    if ($MODE eq 'terminal') {
      print $s,"\n";
    }
    else {
      Msg::put ('IPC.TCPTRACE',$s) ;
    }
  }
}
  
 
sub wait_child {
  # print "Handle child death\n";
  local ($pid) = waitpid (-1,0) ;
}

sub wait_broken_pipe {}

sub install_sigchild {
  $SIG{'CHLD'} = "Tcp::wait_child";
}

sub install_sigpipe {
  $SIG{'PIPE'} = "Tcp::wait_broken_pipe";
}

sub create_service {

  # Description
  #   Create a TCP service.
  #
  # Synopsis
  #   create_service (Out socket, port_number)


  local ($S, $port) = @_ ;
  local ($name, $aliases, $this, $status, $addr, $proto);

  ($name, $aliases, $port) = getservbyport ($port,'tcp') unless $port =~ /^\d+$/ ;

  $this = pack ($sockaddr, $AF_INET, $port, "\0\0\0\0");

  ($name, $aliases, $proto) = getprotobyname ('tcp');
  $status = socket ($S, $AF_INET, $SOCK_STREEM, $proto) ;
  if (! defined $status) {
    trace ("*$$ create_service @_;$!") ;
    return $status ;
  }
  setsockopt ($S, $SOL_SOCKET, $SO_REUSEADDR, 1) ;

  $status = bind ($S, $this) ;
  if (! defined $status) {
    trace ("*$$ create_service @_;$!") ;
    return $status ;
  }

  listen ($S, $MAX_PENDING_CALL) ;

  # Configure the file handle and socket for non buffered input-output
  # operation.

  trace (">$$ create_service @_;$!") ;
  select ((select ($S), $|=1)[$[]) ;

  return 1;  # Successful completion

}


sub create_link {
 
  # Description 
  #   Create a link consisting of a socket pair used for the 
  #   interprocess communication within a given application.

  local ($S1, $S2, $port) = @_ ;
  local ($sts) ;

  $sts = socketpair ($S1, $S2, $AF_UNIX, $SOCK_STREEM, 0) ;
  trace (">$$ create_link @_;$!") ;
  return $sts;
}


sub delete_service {

  # Description
  #   Remove the specified service.
  #
  # Synopsis
  #   delete_service (socket)


  local ($S) = @_ ;
  close $S;
  trace (">$$ close @_;$?") ;
}


sub accept_call {

  # Description
  #   Accept an incoming call  and returned the device calling address.
  #
  # Synopsis
  #   accept_call (Socket, Out Subsocket) return calling_address


  local ($S, $NS) = @_ ; 
  local ($addr, $af, $port, $inetaddr);

  select ((select ($NS), $|=1)[$[]) ;

  if (! defined ($addr = accept ($NS,$S))) {
    trace ("*$$ accept @_;$addr;$!") ;
    return $addr ;
  }

  ($af,$port,$inetaddr) = unpack ($sockaddr,$addr);
  @inetaddr = unpack ('C4',$inetaddr);
  trace (">$$ accept @_;@inetaddr;$!") ;
  return join ('.',@inetaddr) ;
}


sub connect_host {

  # Description
  #   Connect to a remote or local host.
  #
  # Synopsis
  #   connect (Out socket, target address, port)

  local ($S, $them, $port) = @_ ;
  local ($name, $aliases, $this, $that, $status, $proto, $type, $len);
  local ($thisaddr, $hostname) ;
  local ($k);

  while ($hostname eq "") {  ## Due to a possibly BUG, the hostname sometimes
	                     ## return a null string. Needs retry.
			     ## PWI 14-JUL-1999

    chop ($hostname = `hostname`) ;
    if ($k++ > 1) {
      print "Tcp::hostname retry!\n" ;
      sleep 1;
    }
  }

  ## print "### hostname.status = $!, hostname = $hostname\n";

  ($name,$aliases,$proto) = getprotobyname ('tcp');

  ($name,$aliases,$port) = getservbyname ($port,'tcp') unless $port =~ /^\d+$/;
  ($name,$aliases,$type,$len,$thisaddr) = gethostbyname ($hostname);
  ($name,$aliases,$type,$len,$thataddr) = gethostbyname ($them);

  ## Due to a possibly BUG, the localhost address does not work properly.
  ## So, transform the target address by its real host name 
  ## if the target is localhost. (PWI 14-JUL-1999)

  $thataddr=$thisaddr if $thataddr eq $LOCAL;

  $this = pack ($sockaddr, $AF_INET, 0,     $thisaddr);
  $that = pack ($sockaddr, $AF_INET, $port, $thataddr);

  $status = socket ($S, $AF_INET, $SOCK_STREEM, $proto) ;
  if (! defined $status) {
    trace ("*$$ connect.socket @_;$!") ;
    return $status ;
  }

 
  $status = bind ($S, $this) ;
  if (! defined $status) {
    trace ("*$$ connect.bind @_;$!") ;
    return $status ;
  }
 
  $status = connect ($S, $that) ;
  if (! defined $status) {
    trace ("*$$ connect.connect @_;$!") ;
    return $status ;
  }

  select ((select ($S), $|=1)[$[]) ;

  trace (">$$ connect @_;$!") ;
  return 1; # Successful completion

}


sub receive_data {

  # Description
  #   Receive a data on the specified socket.
  #
  # Synopsis
  #   receive (socket, [buf_size=16384]) return data.

  local ($NS, $buf_size, $no_trace) = @_ ;
  local ($message, $res) ;

  $buf_size = $BUF_SIZE unless defined $buf_size;

  $res = recv ($NS, $message, $buf_size, 0) ;
  trace (">$$ receive @_;$message;$!") if $DEBUG && !$no_trace;

  ## printf "message = %s\n", join (':',unpack ('C*',$message)) ;

  return $message ;
}



sub send_data {

  # Description
  #   Send a data to the specified socket and return the number of
  #   bytes that have been sent.
  #
  # Synopsis
  #   send_data (socket, data)


  local ($NS, $message, $no_trace) = @_ ;
  local ($res) ;

  $res = send ($NS, $message, 0) ;
  trace (">$$ send @_;$!") if $DEBUG && !$no_trace;
  return $res;
}


sub close_link {

  # Description
  #   Close the specified socket
  #
  # Synopsis
  #   close_link (socket)


  local ($NS) = @_ ;

  close $NS ;
  trace (">$$ close @_;$?") ;
}

1;

