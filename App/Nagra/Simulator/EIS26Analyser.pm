#!/usr/local/bin/perl

package EIS26Analyser;

require "sms_protocol.pl";


# External: FIELD_DESC
#	    SMS_COMMAND


sub sizeOf {
  local ($field) = @_;
  if ($FIELD_DESC{$field} =~ /^([A-Z]\w*)\s+(\w+)/) {
    return $GfieldValue{$1};
  }
  else {
    $FIELD_DESC{$field} =~ /^(\d+)\s+(\w+)/;
    return $1;
  }
}


sub signalError {
  local ($field, *RESULTS) = @_;
  if ($FIELD_DESC{$field} =~ /(\d+)\s+(\w+)\s+(\w+)\s+(\w+)/) {
    ($fieldSize, $fieldTypeCat, $fieldType, $errorCode) = ($1,$2,$3,$4);
    foreach $code (keys %SMS_ERROR_CODE_EXTENSION) {
      if ($SMS_ERROR_CODE_EXTENSION{$code} eq $errorCode) {
        $errorNr = $code;
        last;
      }
    }
    push @RESULTS, "E $errorCode ($errorNr) for field $field. $fieldType expected";
  }
}


sub checkType {
  local ($field, $value,*RESULTS) = @_;
  if ($FIELD_DESC{$field} =~ /(\d+)\s+(\w+)\s+(\w+)/) {
    ($fieldSize, $fieldTypeCat, $fieldType) = ($1,$2,$3);

    if (length($value) != sizeOf($field)) {
      signalError ($field, *RESULTS);
      return;
    }

    if ($value !~ /^$FIELD_TYPE{$fieldType}$/) {
      signalError ($field, *RESULTS);
      return;
    }

    if ($fieldType eq 'date') {
      $value =~ /^(\d{4,4})(\d{2,2})(\d{2,2})$/ ;
      if (!( $1 >= 1990 && $1 <= 2020 && $2 >= 1 && $2 <= 12 && 
             $3 >= 1 && $3 <= 31)) {
        signalError ($field, *RESULTS);
      }
    }
  }
}


sub getKeyPosition {
  local (*segment) = @_;
  local ($keyField) = $segment{key};
  local ($pos) = 0;
  local ($key, $field);

  foreach $key (keys %segment) {
    next if $key eq 'key';
    local (*selectedSegment) = $segment{$key};
    foreach $field (@selectedSegment) {
      last if $field eq $keyField;
      $pos += sizeOf ($field);
    }
    return $pos;
  }
}


sub getSegment {
  local ($cmd, $segmentName, *RESULTS) = @_;
  local (*segment) = $segmentName;
  local ($keyField) = $segment{key};
  local ($offset) = 0;

  $IDoffset = getKeyPosition($segmentName);
  $id = substr ($cmd, $IDoffset, sizeOf ($keyField));
  $keyId = "\"$id\"";

  push @RESULTS, "D keyId = $keyId";

  local (*selectedSegment) = $segment{$keyId} ;

  if (!defined $segment{$keyId}) {
    push @RESULTS, "E Bad segment identifier ($id) for $segmentName";
    return -1 ;
  }

  foreach $field (@selectedSegment) {
    
    if ($field =~ /^%(\w+)/) {
      $cmd = substr ($cmd, $offset, 9999);
      $offset = getSegment ($cmd, $1, *RESULTS);
      return -1 if $offset == -1;
    }
    else {
      $fieldSize = sizeOf ($field);
      $fieldValue = substr ($cmd, $offset, $fieldSize);
      $GfieldValue{$field} = $fieldValue;
      $offset += $fieldSize;
      checkType ($field, $fieldValue, *RESULTS);
      push @RESULTS, "F $field = $fieldValue";
    }
  }
  return $offset;
}


sub dumpCmd {
  local ($cmd, *RESULTS) = @_;
  $#RESULTS = -1;
  getSegment ($cmd, $SMS_COMMAND[0], *RESULTS);
}

1;
