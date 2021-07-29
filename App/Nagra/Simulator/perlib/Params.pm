
#--+-------------------------------------------------------------------------+--
#--+ package Params
#--+-------------------------------------------------------------------------+--
#--+ Author              : Patrice Willemin                                  +--
#--+ CopyRight           : NAGRAVision, (c) 1994-1998                        +--
#--+ Creation            : 25-JUN-1998                                       +--
#--+ Update              :                                                   +--
#--+ Version             : 1.0                                               +--
#--+ Previous version    : /                                                 +--
#--+                                                                         +--
#--+ Description                                                             +--
#--+      		 : Provide routines to handle parameters
#--+-------------------------------------------------------------------------+-- 
package Params ;

require shared_data;

sub load_module {
  local ($module) = @_ ;
  local ($tmp_file) = "$ENV{'tmp'}/oda_$$.tmp";
  local ($table, $definition, $object, $argument, $index, $value, $key, $skey) ;

  system ("oda $module > $tmp_file") ;

  open (ODA, "$tmp_file") ;
  while (<ODA>) {

    ($table, $definition, $object, $argument, $index) = ('','','','','');

    s/\/(\w*)\/ (\w+)\$//;  ($table, $definition) = ($1,$2);
    s/OBJ\$(\w+)\$// &&   ($object = $1);
    s/(\w+)// &&            ($argument = $1);
    s/\$(\d+)// &&          ($index = $1);
    s/ = \/(.*)\/$//;        $value = $1;

    $value =~ s/^FALSE$/0/;
    $value =~ s/^TRUE$/1/;

    # Substitute $XXX by the value of the corresponding environment variable 

    while ($value =~ s/(.*)\$(\w+)//) {
      $value = $1 . $ENV {$2} . $value ;
    }

    $key = "$table.$definition.$object.$argument.$index";


    if (shared_data::controller_started()) {
       shared_data::set ("Params.DEFINED.$key", $value);
    }
    else {
      $DEFINED {"$table.$definition.$object.$argument.$index"} = $value; 
    }
  }
  unlink $tmp_file;
}


sub display {
  (%ARR) = @_ ;
  foreach $name (sort keys (%ARR)) {
    printf "%-40s%-20s\n", $name, $ARR {$name};
  }
}


sub list {
  local ($table) = @_;		# Table = USED | DEFINED
  local (%RES) ;

  if ($table ne 'USED' && $table ne 'DEFINED') {
    print "Params.param: invalid argument $table\n";
    return ;
  }

  if (shared_data::controller_started()) {
    local (%PM_RES) ;
    local (%PM) = shared_data::select_table ("^Params.$table") ;
    local ($skey,$key);

    foreach $skey (keys %PM) {
      $key = $skey;
      $key =~ s/^Params.$table.//;
      $PM_RES{$key} = $PM{$skey};
    }
    return %PM_RES;
  }
  else {
    eval "\%RES = \%$table";
    return %RES ;
  }
}



sub val {
  local ($key) = @_;
  local ($value) ;
  local ($skey);

  if (shared_data::controller_started()) {
    $value = shared_data::get ("Params.DEFINED.$key") ;
  }
  else {
    $value = $DEFINED {$key} ;
  }

  if (!defined $value) {
    print "%PM-E-UNDEF, Undefined parameter $key\n";
  }
  else {
    if (shared_data::controller_started()) {
       $skey = "Params.USED.$key";
       shared_data::set ($skey, shared_data::get ($skey)+1); 
    }
    else {
      $USED {$key}++;
    }
    return $value ;
  }
}
1
