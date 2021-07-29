

#--+-------------------------------------------------------------------------+--
#--+ package Trace
#--+-------------------------------------------------------------------------+--
#--+ Author              : Patrice Willemin                                  +--
#--+ CopyRight           : NAGRAVision, (c) 1994-1998                        +--
#--+ Creation            : 16-JUN-1998                                       +--
#--+ Update              :                                                   +--
#--+ Version             : 1.0                                               +--
#--+ Previous version    : /                                                 +--
#--+                                                                         +--
#--+ Description                                                             +--
#--+                     : Provide trace handling procedure
#--+-------------------------------------------------------------------------+--

package Trace;

require Time ;

sub put {
  local ($text) = @_ ;
  local (@callinfo,$callid);

  print STDOUT Time::image(time)," $$ ",$text,"\n";
  while (@callinfo = caller($callid++)) {
    print "                     $callid", "<" . join (':',@callinfo),"\n";
  }
} 
1;
