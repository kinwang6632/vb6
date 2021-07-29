
#--+-------------------------------------------------------------------------+--
#--+ package Terminal
#--+-------------------------------------------------------------------------+--
#--+ Author              : Patrice Willemin                                  +--
#--+ CopyRight           : NAGRAVision, (c) 1994-1998                        +--
#--+ Creation            : 03-JUN-1998                                       +--
#--+ Update              :                                                   +--
#--+ Version             : 1.0                                               +--
#--+ Previous version    : /                                                 +--
#--+                                                                         +--
#--+ Description                                                             +--
#--+                     : Provide primitives for handling terminal
#--+			   operations.
#--+-------------------------------------------------------------------------+--

package Terminal;

$CSI	= "\033[" ;

sub at {

  # Description
  #   Set the cursor position on the current terminal

  ($line, $column) = @_;

  print "${CSI}${line};${column}H" ;
}

sub set_bold {
  print "${CSI}1m";
}

sub reset_bold {
  print "${CSI}22m";
}

sub cls {
  
  # Description 
  #   Clear the terminal screen
  #
  # Synopsis 
  #   cls ([mode = (full|bottom|top)])

  local ($mode) = @_ ;
  $mode = 'full' if !defined $mode ;

  CASE : {
    $mode =~ /full/   && print "${CSI}2J";
    $mode =~ /bottom/ && print "${CSI}0J";
    $mode =~ /top/    && print "${CSI}1J";
  }
}
1; 
