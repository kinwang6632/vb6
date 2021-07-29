
#--+-------------------------------------------------------------------------+--
#--+ package File_mgr
#--+-------------------------------------------------------------------------+--
#--+ Author              : Patrice Willemin                                  +--
#--+ CopyRight           : NAGRAVision, (c) 1994-1998                        +--
#--+ Creation            : 03-JUN-1998                                       +--
#--+ Update              :                                                   +--
#--+ Version             : 1.0                                               +--
#--+ Previous version    : /                                                 +--
#--+                                                                         +--
#--+ Description                                                             +--
#--+                     : Provide File management routines.
#--+-------------------------------------------------------------------------+--

package File_mgr ;

$ATIME = 8;	# Last access time
$MTIME = 9;	# Last modification time
$CTIME = 10 ;	# Last status change

sub last_version {

  # Description
  #   Return the highest version of the given file located in the
  #   specified directory.
  #
  # Synopsis
  #   last_version (directory_name, file_name)


  local ($dir_name,$file_body) = @_;
  local ($max_ver) = 0;
  local ($file);

  opendir (DIRECTORY, $dir_name) ;
  while (1) {
    $file = readdir (DIRECTORY);
    last unless $file;
    if ($file =~ /^$file_body-(\d+)/) {
      if ($1 > $max_ver) {$max_ver = $1}
    }
  }
  closedir (DIRECTORY) ;
  $max_ver;
}


sub line_count {

  # Description
  #   Return the number of lines in the given file matching a specified
  #   character string pattern.
  #
  # Synopsis
  #   line_count (file_name, [filter])


  local ($file,$filter) = @_;
  local ($N) = 0;

  open (FILE,$file) ;
  while (<FILE>) {$N++ if /$filter/};
  close (FILE) ;
  $N;
}

sub get_files_by_date {
  # Description
  #   Get the list of the file in the specified directory matching the
  #   given selection. The files are sorted by date.
  #
  # Synopsis
  #   get_files_by_date (directory, Out File_list, [selection], since,until)

  local ($dir_name, *file_list, $selection, $since, $until) = @_;
  local ($p) = 0;
  local ($file,@full_spec,@cp_file_list,$mtime);
  $since = 0 if !defined $since;
  $until = 0 if !defined $until;

  if (substr ($selection,length($selection)-1,1) eq '$') {
    chop ($selection);
  }
  else {
    $selection = $selection . ".*" ;
  }

  @file_list=();
  opendir (DIRECTORY, $dir_name) ;
  while (1) {
    $file = readdir (DIRECTORY);
    last unless $file;
    if ($file =~ /$selection/) {
      local (@file_attr) = stat ("$dir_name/$file");
      $mtime = $file_attr[$MTIME];
      push @file_list, sprintf ("%10.10d",$mtime).' '.$file if
        $mtime > $since && ($mtime < $until || $until == 0) ;
    }
  };
  closedir (DIRECTORY) ;

  @file_list = sort @file_list;
  foreach (@file_list) {
    s/(\d+) //;
  }
}

sub get_files_by_versions {

  # Description
  #   Get the list of the file in the current directory matching the
  #   given selection. The file are sorted according to the name and
  #   to the current file version (i.e. FILE-<ver>.)
  #   Only the files ending by a version number are selected.
  #
  # Synopsis
  #   get_files_by_versions (directory, Out File_list, [selection])


  local ($dir_name, *file_list, $selection) = @_;
  local ($p) = 0;
  local ($file,@full_spec,@cp_file_list);

  if (substr ($selection,length($selection)-1,1) eq '$') {
    chop ($selection);
  }
  else {
    $selection = $selection . ".*" ;
  }

  @file_list=();
  opendir (DIRECTORY, $dir_name) ;
  while (1) {
    $file = readdir (DIRECTORY);
    last unless $file;
    if ($file =~ /(${selection})-(\d+)$/) {
      # transform version number in 10 decimal digits
      push (@file_list, "$1-" . sprintf ("%10.10d",$2));
    }
  };
  closedir (DIRECTORY) ;

  @file_list = sort @file_list;
  foreach (@file_list) {
    s/(.*)-[0]*([1-9]*)/\1-\2/; # Suppress heading 0 in version number
    s/(.*)-$/\1-0/;		# Correction for file with null version.
  }
}


sub purge {

  # Description
  #   Purge the N last versions of the select files.
  #   The file version is located at the end of the file name (file_<ver>).
  # Synopsis
  #   purge (directory, $selection, [keep_max=1])
  
  local ($dir_name, $selection, $keep_max) = @_;
  local (@file_list, $last_file, $count);

  $keep_max=1 if $keep_max eq "" ;

  get_files_by_versions ($dir_name, *file_list, $selection) ;
  foreach (reverse @file_list) {
    /(.*)-(\d+)$/;
    $count++;
    $count = 0 if $1 ne $last_file ;
    unlink ("$dir_name/$_") 
      if $1 eq $last_file && $count >= $keep_max || $keep_max == 0 ;
    $last_file = $1;
  }
}

sub components {

  # Description
  #   return the file directory and the file body according to a
  #   full file specification

  local ($full_spec) = @_;
  if ($full_spec =~  /(.*)\/([^\/]+)$/) { 
    return ($1,$2);
  }
  else {
    return ("",$full_spec) ;
  }
}


sub fetch_files_from_dir {
  local ($root,*file_list) = @_ ;
  local (@files, $file) ;

  opendir (DIRHANDLE, $root);
  @files = readdir DIRHANDLE;
  closedir DIRHANDLE ;

  foreach $file (@files) {
    next if $file =~ /^\./;
    push (@file_list, "$root/$file");
    fetch_files_from_dir ("$root/$file",*file_list) if -d "$root/$file";
  }
}

sub all_files {
  local ($root) = @_ ;
  local (@file_list) ;

  fetch_files_from_dir ($root, *file_list) ;
  return @file_list;
}

1;
