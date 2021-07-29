
#--+-------------------------------------------------------------------------+--
#--+ package Ipc
#--+-------------------------------------------------------------------------+--
#--+ Author              : Patrice Willemin                                  +--
#--+ CopyRight           : NAGRAVision, (c) 1994-1998                        +--
#--+ Creation            : 03-JUN-1998                                       +--
#--+ Update              :                                                   +--
#--+ Version             : 1.0                                               +--
#--+ Previous version    : /                                                 +--
#--+                                                                         +--
#--+ Description                                                             +--
#--+                     : Provide routine for inter-process communication
#--+-------------------------------------------------------------------------+--

package Ipc;


$IPC_CREAT      = 0x0001000;
$IPC_EXCL       = 0x0002000;
$IPC_NOWAIT     = 0x0004000;

$IPC_PRIVATE    = 0;

$IPC_RMID       = 0;
$IPC_SET        = 1;
$IPC_STAT       = 2;

$GETVAL         = 5;
$SETVAL         = 8;

$EINTR		= 4;	# Interrupted system call

sub getperm {

  # Description
  #   INTERNAL usage ONLY
  #   Set the permission data structure.


  # Permission data structure
  #
  local ($UID,$GID) = 		# User and Group id. of the current process.

    split (' ', `ps -p $$ -eo svuid, svgid | tail -1`);


  local ($CUID)	= $UID;		# Creator's user id.
  local ($CGID) = $GID;		# Creator's group id
  local ($MODE) = 		# Permission flags (0777).

    pack ('C4', 0xFF, 0x81, 0x00, 0x00);


  local ($SEQ)	= 0;		# Slot usage sequence number.	
  local ($KEY)  = 1;		# key.

  return pack ('I4 a4 I2', $UID, $GID, $CUID, $CGID, $MODE, $SEQ, $KEY);

}


sub shm_delete {

  # Description
  #   Remove the shared memory segment.
  #
  # Synopsis
  #   delete segment_id

  local ($shmid) = @_ ;
  local ($status);

  $status = shmctl ($shmid, $IPC_RMID, 0) ;
  if (! defined $status ) {
    return $status ;
  }
}


sub shm_create {

  # Description
  #   Create a shared memory segment of the specified size.
  #
  # Synopsis
  #   create (segment_size) return segment_identifier.


  local ($SEGSZ) = @_ ;		# Segment size

  # Shared memory data structure
  #
  local ($PERM) = getperm(); 	# IPC permission data structure.
  local ($LPID) = $$ ;		# Pid if the last shm operation.
  local ($CPID) = $$ ;		# Pid of the creator.
  local ($NATTCH_H) = 0;	# current number attached.
  local ($NATTCH_L) = 0;	#   
  local ($ATIME) = time;	# last shmat (shm attach) time.
  local ($DTIME) = time;	# last shmdt (shm detach) time.
  local ($CTIME) = time;	# last change time.
  local ($SECATTR) = 0;		# Security attributes pointer.


  local ($SHMID) = pack ('a28 I9', $PERM, $SEGSZ, $LPID, $CPID, $NATTCH_L,
                        $NATTCH_H, $ATIME, $DTIME, $CTIME, $SECATTR) ;

  local ($shmid, $status) ;
    

  $shmid  = shmget ($IPC_PRIVATE, $SEGSZ, $IPC_CREAT) ; 
  if (! defined $shmid ) {
    return $shmid ;
  }

  $status = shmctl ($shmid, $IPC_SET, $SHMID) ;
  if (! defined $status ) {
    shmctl ($shmid, $IPC_RMID, 0) ;
    return $status;
  }

  return $shmid;

}   


sub sem_delete {

  # Description
  #   Remove the semaphore set
  #
  # Synopsis
  #   delete semid

  local ($semid) = @_ ;
  local ($status);

  $status = semctl ($semid, 0, $IPC_RMID, 0) ;
  if (! defined $status ) {
    return $status ;
  }
}


sub sem_create {

  # Description
  #   Create a semaphore set and set an initial value for each semaphore
  #   in the set.
  #
  # Synopsis
  #   create ([semaphore_number],[init_value]) return semaphore_identifier.


  local ($NSEMS,$VAL) = @_ ;		

  $NSEMS=1 if !defined $NSEMS;
  $VAL=1 if !defined $VAL;

  # Semaphore data structure
  #
  local ($PERM) = getperm(); 	# IPC permission data structure.
  local ($SEM_BASE_L) = 0;	# ptr to first semaphore in set.
  local ($SEM_BASE_H) = 0;	
  local ($OTIME) = time ;	# last semop time.
  local ($CTIME) = time ;	# last change time.
  local ($SECATTR) = 0;		# Security attributes pointer.


  local ($SEMID) = pack ('a28 I7', $PERM, $SEM_BASE_L, $SEM_BASE_H, $NSEMS,
                        $OTIME, $CTIME, $SECATTR, 0) ;

  local ($semid, $status, $k) ;
    

  $semid  = semget ($IPC_PRIVATE, $NSEMS, $IPC_CREAT) ; 
  if (! defined $semid ) {
    return $semid ;
  }

  $status = semctl ($semid, 1, $IPC_SET, $SEMID) ;
  if (! defined $status ) {
    semctl ($semid, 0, $IPC_RMID, 0) ;
    return $status;
  }

  for ($k=0; $k<$NSEMS; $k++) {
    semctl ($semid, $k, $SETVAL, $VAL) ;
  }

  return $semid;

}   


sub sem_f_create {

  # Description
  #   Create semaphore. If, after a crash, the semaphore id file is
  #   present, the semaphore, of which the id is given by this file is
  #   first removed.
  #
  # Synopsis
  #   sem_f_create (sem file, [nb semaph=1], [semph initial value=1])

  local ($sem_file, $NSEMS, $VAL) = @_ ;	
  return 0 if !defined $sem_file ;

  if (-r $sem_file) {
    open (FILE, $sem_file) ;
    $SEMID = <FILE>;
    Ipc::sem_delete ($SEMID) ;
    close FILE;
  }
  open (FILE, "> $sem_file") ;   # Store the new semaphore id
  $SEMID = sem_create ($NSEMS, $VAL) ;
  print FILE $SEMID;
  close FILE ;
  return $SEMID ;
}


sub sem_f_delete {
  # Description
  #   Delete the file and the specified attached semaphore.
  #
  # Synopsis
  #   sem_f_delete (sem file)

  local ($sem_file) = @_ ;
  local ($SEMID) ;

  if (-r $sem_file) {
    open (FILE, $sem_file) ;
    $SEMID = <FILE>;
    Ipc::sem_delete ($SEMID) ;
    close FILE;
    unlink $sem_file;
    return 1;
  }
  return 0;
}


sub sem_post {

  # Description
  #   Post the given semaphore
  #
  # Synopsis
  #   sem_post (semset_id, [sem_nr=0])


  local ($semid, $semnr) = @_ ;
  $semnr = 0 if !defined $semnr;

  semop ($semid, pack ("s*",$semnr,1,0)) ;
}


sub sem_wait {

  # Description
  #   Wait on the given semaphore
  #
  # Synopsis
  #   sem_wait (semset_id, [sem_nr=0])


  local ($semid, $semnr) = @_ ;
  local ($sts, $callid, @callinfo) ;
  $semnr = 0 if !defined $semnr;

  while (1) {
    $sts = semop ($semid, pack ("s*",$semnr,-1,0)) ;
    last if $! != $EINTR ;

    # Remove comment for debug.
    # print STDOUT Time::image (time)," SEM $semid;$!\n" if $! ;
  }
}

sub sem_val {

  # Description
  #   Return the semaphore value
  #
  # Synopsis
  #   sem_val (semset_id, [sem_nr=0]) return value.

  local ($semid, $semnr) = @_ ;
  $semnr = 0 if !defined $semnr;

  return semctl ($semid, $semnr, $GETVAL, 0) ;
}

sub sem_set {

  # Description
  #   Set the semaphore with the given value
  #
  # Synopsis
  #   sem_set (semset_id, sem_nr, sem_value)

  local ($semid, $semnr, $value) = @_ ;

  semctl ($semid, $semnr, $SETVAL, $value) ;
}

1;
