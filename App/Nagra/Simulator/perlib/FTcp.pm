

#--+-------------------------------------------------------------------------+--
#--+ package FTcp
#--+-------------------------------------------------------------------------+--
#--+ Author              : Patrice Willemin                                  +--
#--+ CopyRight           : NAGRAVision, (c) 1994-1998                        +--
#--+ Creation            : 03-JUN-1998                                       +--
#--+ Update              :                                                   +--
#--+ Version             : 1.0                                               +--
#--+ Previous version    : /                                                 +--
#--+                                                                         +--
#--+ Description                                                             +--
#--+                     : Provide FTcp communication primitives
#--+-------------------------------------------------------------------------+--

  
package FTcp;

# Package summary
#
# -- create_service (Out socket, port number)
# -- delete_service (socket)
# -- accept_call (socket, Out subsocket)
# -- connect_host (Out socket, target address, port number)
# -- receive_data (socket) return data
# -- send_data (socket, data) return nb.bytes sent
# -- close_link (socket)

 
require Tcp;
require Ipc;
require Time;
require Trace;
require Msg;


sub install_sigchild {
  Tcp::install_sigchild();
}

sub install_sigpipe {
  Tcp::install_sigpipe();
}


# Global variables
#
$DEBUG = 0 ;


sub set_trace_mode {

  # Description
  #   Set the trace mode for FTCP communication.
  #
  # Synopsis
  #   set_trace_mode (trace_file_name, [semaphore id], [semaphore number])

  $DEBUG = 1;
}

sub trace {
  local ($text) = @_ ;
  local ($sts) = 1 ;

  Msg::put ('IPC.FTCPTRACE',$text) if $DEBUG;
}


sub create_service {

  # Description
  #   Create a FTCP service.
  #
  # Synopsis
  #   create_service (Out socket, port_number)

  local ($sts) ;

  trace ("<$$ create @_") ;
  $sts = Tcp::create_service (@_) ;
  trace (">$$ create @_;$!") ;
  return $sts;
}


sub create_link {

  # Description
  #   Create a link consisting of a socket pair used for the
  #   interprocess communication within a given application.
  #
  # Synopsis
  #   create_link (Out Socket1, Out Socket2)

  local ($sts) ;

  trace ("<$$ create_link @_") ;
  $sts = Tcp::create_link (@_) ;
  trace (">$$ create_link @_;$!") ;
  return $sts;
}
    
 
sub delete_service {

  # Description
  #   Remove the specified service.
  #
  # Synopsis
  #   delete_service (socket)

  local ($sts);
 
  trace ("<$$ delete @_") ; 
  $sts = Tcp::delete_service (@_) ;
  trace (">$$ delete @_;$!") ;
  return $sts ;
}


sub accept_call {

  # Description
  #   Accept an incoming call  and returned the device calling address.
  #
  # Synopsis
  #   accept_call (Socket, Out Subsocket) return calling_address


  local ($S, $NS) = @_ ; 
  local ($inetaddr,$block,$sts) ;

  trace ("<$$ accept @_") ;
  $inetaddr = Tcp::accept_call (@_) ;
 
  $block = receive_data ($NS) ;
  send_data ($NS, "\6") ;
  send_data ($NS, "\0") ;
  trace (">$$ accept @_;$inetaddr;$!") ;
  return $inetaddr;
}


sub connect_host {

  # Description
  #   Connect to a remote or local host.
  #
  # Synopsis
  #   connect (Out socket, target address, port)

  local ($sts) ;
  local ($S, $target, $port) = @_;

  trace ("<$$ connect @_") ;

  $sts = Tcp::connect_host (@_) ;
  if (!defined $sts) {
    trace ("*$$ connect @_;$!");
    return $sts ;
  }
  $sts=Tcp::send_data ($S, pack ('C4',0,2,0,0));  
  receive_data ($S) ;
  receive_data ($S) ;

  trace (">$$ connect @_;$!") ;
  return 1; # Successful completion

}


sub receive_data {

  # Description
  #   Receive a data on the specified socket.
  #
  # Synopsis
  #   receive (socket, [buf_size=16384]) return data.

  local ($NS,$no_trace) = @_ ;
  local ($res, $len, $head, $block,$sts) ;
  local ($data) = "";
  local ($undefined);

  trace ("<$$ receive @_") if $DEBUG && !$no_trace ;

  $head = Tcp::receive_data ($NS, 2, $no_trace);
  if ($head eq '') {
    trace ("*$$ receive @_;$!") if !$no_trace;
    return $undefined ;
  }

  $head .= Tcp::receive_data ($NS, 1, $no_trace) if length($head) < 2;
  ($len)  = unpack ('n',$head) ;
  while (length ($data) < $len) {
    $data .= ($block = Tcp::receive_data ($NS, $len-length($data),$no_trace));
    if ($block eq '') {
      trace ("*$$ receive @_;$!") if !$no_trace;
      return $undefined ;
    }
  }
  trace (">$$ receive @_;$data;$!") if $DEBUG && !$no_trace;
  return $data;
}



sub send_data {

  # Description
  #   Send a data to the specified socket and return the number of
  #   bytes that have been sent.
  #
  # Synopsis
  #   send_data (socket, data)


  local ($NS, $message, $no_trace) = @_ ;
  local ($sts) ;

  trace ("<$$ send @_") if $DEBUG && !$no_trace;
  $sts = Tcp::send_data ($NS, pack ('n a*',length ($message), $message),
         $no_trace);
  trace (">$$ send @_;$!") if $DEBUG && !$no_trace;
  return $sts;

}


sub close_link {

  # Description
  #   Close the specified socket
  #
  # Synopsis
  #   close_link (socket)

  local ($sts) ;

  trace ("<$$ close @_") ;
  $sts = Tcp::close_link (@_) ;
  trace (">$$ close @_;$!") ;
  return $sts;
  
}

1;

