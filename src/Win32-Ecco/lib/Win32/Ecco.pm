#!/usr/perl/bin/perl 

# A package that exposes Ecco's DDE API as a set of functions.

# FIXME Need MUCH better documentation
# FIXME Need croak() assertions on all methods

# Email will_sargent@yahoo.com

package Win32::Ecco;

use 5.008004;
use strict;
use warnings;
use Carp;

use Date::Parse;
use Date::Format;

use Win32::DDE;
use Win32::DDE::Client;

use Exporter;
use vars qw($VERSION @ISA @EXPORT);

$VERSION=1.00;
@ISA=qw(Exporter);
@EXPORT=
    qw(&filterItem 
    &getFolderId 
    &createFolder 
    &createItem 
    &insertItem
    &getFolderType
    &getFolderValues
    &setFolderValues
    &getFolderItems
    &getFolderName
    &setFolderName
    &setItemText
    &getItemSubs
    &getViews
    &getViewNames
    &getViewFolders
    &getViewTLIs
    &getViewsHash
    &dequote
    &getCurrentFile
    &getFileName
    &getOpenFiles
    &getPencilInDateFolderValues
    &getColorDateFolderValues
    &getTicklerDateFolderValues
    &getAlarmDateFolderValues
    &getRecurringDateFolderValues
    &getRecurDateFolderValues
    &getExtendedDateFolderValues
    &processStartDate
    &processEndDate
    &getDateValues
    &formatDate);

my %FOLDER_TYPES = (
	checkmark => 1,
	date => 2,
	number => 3,
	text => 4,
	popup => 5,
);

my @EXTENDED_DATE_TYPES = ( "PencilIn", "Tickler", "Color", "Alarm", "Recur" );

my $DEBUG = 0;

sub new()
{
    my $self = {};
    $self->{CONNECTION} = new Win32::DDE::Client("Ecco", "Ecco");
    die "Error connecting to Ecco!" if ($self->{CONNECTION}->Error);
    bless $self;
    return $self;
}

sub poke($$)
{
    my ($self, $cmd) = @_;
    my $conn = $self->{CONNECTION};
    my $pokeValue = $conn->Poke($cmd);
    if ($conn->Error) 
    {
       my $errorText = Win32::DDE::ErrorText($conn->Error);
       die "DDE request failed on [$cmd], errorText = $errorText";   
    }
}

sub request($$)
{    
   my ($self, $cmd) = @_;       
   my $conn = $self->{CONNECTION};
   my $requestValue = $conn->Request($cmd);
   
   if ($conn->Error) 
   {
       my $errorText = Win32::DDE::ErrorText($conn->Error);
       die "DDE request failed on [$cmd], errorText = $errorText";   
   }
   
   return $requestValue;    
}

# Filters the item to make sure the folder values of the 
# item meet the criteria we specified.
sub filterItem
{
    my $self = shift;
    my $itemId = shift;
    my %filter = @_;

    print "filterItem: $itemId, %filter\n" if ($DEBUG);

    my @folderIds = keys %filter;
    
    my @folderValues = $self->getFolderValues($itemId, @folderIds);
    
    my $i = 0;
    foreach my $folderId (@folderIds) 
    {
        my $desiredValue = $filter{$folderId};
        my $actualValue = $folderValues[$i++];
            
        if (! defined($desiredValue) && (! defined($actualValue) || $actualValue eq ""))
        {            
            next;    
        } elsif ((! defined($desiredValue)) && defined($actualValue))
        {         
            # print ("$itemId failed: $folderId=>undef != $actualValue\n");
            return 0;
        } elsif ((defined($desiredValue) && (! defined($actualValue))))
        {
            # print ("$itemId failed: $folderId=>$desiredValue != undef\n");        
            return 0;
        } elsif ($desiredValue ne $actualValue)
        {
            # print ("$itemId failed: $folderId=>$desiredValue != $actualValue\n");        
            return 0;
        }
    }
    
    # Otherwise return success.
    # print ("$itemId passed filter with values @folderValues\n");
    return 1;
}

# Gets a new folder id in Ecco.
# Hacked this to return undef if there is no result.
sub getFolderId
{ 
    my $self = shift;
    my $folderName = shift;
    
    print "getFolderId: $folderName\n" if ($DEBUG);

    my $folderId = request($self,"GetFoldersByName,$folderName");           
    
    if ($folderId eq '') 
    {
        return undef;
    }
    
    return $folderId;
}

# Creates a new folder in Ecco.  Only takes one folder type / text.
sub createFolder
{
    my $self = shift;
    my ($folderType, $text) = @_;
    
    # prevent a complaint if the line is blank.
    die "Undefined folder name: $text\n" if (! defined($text));
    
    my $folderId = request($self,"CreateFolder,$folderType,$text");

    return $folderId;
}

# Creates a new item in Ecco
# Accepts: the folderId, the line
# Returns the id of the newly created item.
sub createItem
{
    my $self = shift;
    my ($folderId, $line) = @_;
    
    # prevent a complaint if the line is blank.
    $line = "" if (! defined($line));
    
    my $itemId = request($self,"CreateItem,$line,$folderId");

    return $itemId;
}

# Inserts the child id as a daughter of the parent id.
# parentId the parent id 
# flag the ecco flag indicating how this will be added.
# childId the id of the item to be a child.
sub insertItem
{
    my $self = shift;
    my ($parentId, $flag, $childId) = @_;
        
    poke($self,"InsertItem,$parentId,$flag,$childId");    
}

# Returns a list of folders that the item exists in.
sub getItemFolders
{
    my $self = shift;
    my ($itemId) = @_;

    my $folderList = request($self,"GetItemFolders,$itemId");
    my @folderIds = split(',', $folderList);    
    
    return @folderIds;
}

# Gets the item text specified by $itemId
sub getItemText
{
    my $self = shift;
    my ($itemId) = @_;
    
    print "getItemText: $itemId\n" if ($DEBUG);

    my $itemText = request($self,"GetItemText,$itemId");
    
    return $itemText; 
}

# Takes in a single ids and returns a folder type.
sub getFolderType
{
    my $self = shift;
    my $folderIds = shift;
    
    print "getFolderType: $folderIds\n" if ($DEBUG);

    my $folderTypes = request($self,"GetFolderType,$folderIds");
    my @types = split(/,/, $folderTypes);
    
    if ($#types == -1) 
    {
        return undef;
    }
    return $types[0]; 
}

# Gets the values of $itemId in the given folders.
sub getFolderValues
{
    my $self = shift;
    my ($itemId, @folderIds) = @_;   
        
    print "getFolderValues: $itemId, @folderIds\n" if ($DEBUG);

    croak("getFolderValues: No value defined for itemId") unless (defined($itemId));  
    
    my $folders = join(",", @folderIds);        
    
    my $cmd = "GetFolderValues,$itemId\r$folders";
    # print ("DEBUG getFolderValues: $cmd\n");
        
    my $items = request($self,$cmd); 
    my @itemList = split(/,/, $items, ($#folderIds+1));
    # print ("DEBUG getFolderValues RETURN: [$items] -> [@itemList] = $#itemList\n");
    
    return @itemList;    
}

# NOTE: THE ECCO DOC IS WRONG HERE, IT IS ITEM ID FIRST >:-(
# SetFolderValues < ItemID * > < FolderID * > < FolderValue * > *
# For each item, in the order given, there is a line of folder values showing the new value for each folder, in the order given.
# A null string means that the item will be removed from the folder.
# If a FolderID or ItemID is invalid, then the values in the corresponding column or row will be ignored.
# If a value is invalid, then the item's folder value will not be changed.
sub setFolderValues
{
    my $self = shift;
    my ($itemId, $folderId, $folderValue) = @_;
    
    print "setFolderValues: $itemId, $folderId, $folderValue\n" if ($DEBUG);

    croak("setFolderValues: No value defined for folderId") unless (defined($folderId));  
    
    my $cmd = "SetFolderValues,$itemId\r$folderId\r$folderValue";
    poke($self,$cmd);
}

# Gets the items in the folder specified by $folderId
sub getFolderItems
{
    my $self = shift;
    my ($folderId) = @_;    
    
    croak("getFolderItems: No value defined for folderId") unless (defined($folderId));  
    
    print "getFolderItems: $folderId\n" if ($DEBUG);

    my $itemIds = request($self,"GetFolderItems,$folderId");
    my @itemIdList = split(',', $itemIds);
    return @itemIdList;    
}

# Gets the name of $folderId
sub getFolderName
{
    my $self = shift;
    my ($folderId) = @_;    
    my $folderName = request($self,"GetFolderName,$folderId");    
    return $folderName;
}

# Returns the list of item ids that are the parents of $itemId
sub getItemParents
{
    my $self = shift;
    my ($itemId) = @_;    
    my $parentIds = request($self,"GetItemParents,$itemId");    
    my @parentIdList = split(',', $parentIds);    
    return @parentIdList;  
}

# Sets the folder specified by $folderId to $text.
sub setFolderName
{
    my $self = shift;
    my ($folderId, $text) = @_;    
    poke($self,"SetFolderName,$folderId,\"$text\"");    
}

# Sets the item specified by $itemId to $text.
sub setItemText
{
    my $self = shift;
    my ($itemId, $text) = @_;
    poke($self,"SetItemText,$itemId,\"$text\"");     
}

# Gets the sub items below the current item. 
# Returns an array where each element is the indent level minus one, i.e.
# $items[0] returns an array of the itemIds that are one level below the 
# specified item.
sub getItemSubs
{
    my $self = shift;
    my ($depth, @itemIdList) = @_;

    my $itemIds = join(',', @itemIdList);    
    my $subItems = request($self,"GetItemSubs,$depth,$itemIds");
    
    my @data = split(/,/, $subItems);    
    my @itemTree;    
    while ($#data > 0) 
    {
        my $indentLevel = shift @data;
        my $itemId = shift @data;
	
        # Push itemId onto the array in $itemTree[]
        push @{ $itemTree[$indentLevel - 1] }, $itemId;  
    }
        
    return @itemTree;
}

# Returns a list of views.  This is part of the extended API.
sub getViews
{
    my $self = shift;
    
    my $viewIds = request($self,"GetViews");    
    # print "viewIds = $viewIds\n";
    
    my @viewList = split(/,/, $viewIds);
    
    return @viewList;
}

# Gets the names of the given viewIds.
sub getViewNames
{
    my $self = shift;
    my @viewIdList = @_;
    
    my $viewIds = join(",", @viewIdList);
    
    my $viewNames = request($self,"GetViewNames,$viewIds");    
    # print "viewNames = $viewNames\n";
    my @viewNameList = split(/,/, $viewNames);
    
    return @viewNameList;
}

# Gets the folders available in the view.
# GetViewFolders ViewID * -> < folderID * > *
sub getViewFolders
{
    my $self = shift;
    my @viewIdList = @_;
    
    my $viewIds = join(",", @viewIdList);
    
    my $folderIds = request($self,"GetViewFolders,$viewIds");    
    
    # print "folderIds = $folderIds\n";
    
    my @folderIdList = split(/,/, $folderIds);
    
    return @folderIdList;    
}

# GetViewTLIs ViewID -> < folderID, ( ItemID ) * > *
# Returns a multi-line list of Top Level Item's (TLIs). The first item in each line is          
# the containing folder. All subsequent items on a given line are Top Level Items.
sub getViewTLIs
{
    my $self = shift;
    my $viewId = shift;
    
    if (! defined($viewId))
    {
        die "getViewTLIs: viewId is not defined\n";    
    }
    
    my $response = request($self,"GetViewTLIs,$viewId");     
    
    # print "response = $response\n";
    
    # splits on whitespace, not what you'd expect
    my @lines = split(' ', $response);
    my @folders;
    for my $line (@lines)
    {
        # print "line = $line\n";
        push @folders, [ split(/,/, $line) ];
    }   
    return @folders;    
}

# Returns a hash containing the names of the views as the keys
# and the ids of the views as the values.
sub getViewsHash 
{
    my $self = shift;
        
    my @viewIdsList = $self->getViews();    
    my @viewNames = $self->getViewNames(@viewIdsList);
    
    my %views;
    foreach my $viewId (@viewIdsList) {
        my $viewName = shift @viewNames;
        $viewName = $self->dequote($viewName);
        $views{$viewName} = $viewId;
    }
    
    return %views;
}

# Strips out the quotes from the text.
sub dequote
{    
    my $self = shift;
    
    my $text = shift;

    $text =~ s/"(.*)"/$1/;
    
    # If there are quotes in the actual text, then they're replaced by two quotes.
    # So we reverse the process here.
    $text =~ s/""/"/g;

    return $text;    
}

# Gets the current file.
sub getCurrentFile
{
    my $self = shift;
       
    my $sessionId = request($self,"GetCurrentFile");

    return $sessionId;    
}

# Gets a session id, returns the file name.
sub getFileName 
{
    my $self = shift;
    my $sessionId = shift;
        
    my $filename = request($self,"GetFileName,$sessionId");    
    
    return $filename;       
}

# returns the open files as a list of session ids.
sub getOpenFiles
{
     my $self = shift;
     
     my $sessionIds = request($self,"GetOpenFiles");    
    
     my @sessionIdList = split(/,/, $sessionIds);
     
     return @sessionIdList;
}

################################################################################
# Date information
################################################################################

# PencilIn - "keep time available" setting
# Color - busy bar color
# Tickler - tickler settings
# Alarm - alarm
# Recur - recurrence pattern and exceptions

sub getPencilInDateFolderValues
{
    my $self = shift;
    my $extendedType = shift;
    my ($itemId, @folderIds) = @_;   
    
   return $self->getExtendedDateFolderValues("PencilIn", $itemId, @folderIds);
}

sub getColorDateFolderValues
{
    my $self = shift;
    my $extendedType = shift;
    my ($itemId, @folderIds) = @_;   
    
   return $self->getExtendedDateFolderValues("Color", $itemId, @folderIds);
}

sub getTicklerDateFolderValues
{
    my $self = shift;
    my $extendedType = shift;
    my ($itemId, @folderIds) = @_;   
    
   return $self->getExtendedDateFolderValues("Tickler", $itemId, @folderIds);
}

sub getAlarmDateFolderValues
{
    my $self = shift;
    my $extendedType = shift;
    my ($itemId, @folderIds) = @_;   
    
   return $self->getExtendedDateFolderValues("Alarm", $itemId, @folderIds);
}


# Returns out items with recurring date values in the format:
# "<notes>;<type>;<end>;<interval>;<pattern>[;<except>]*"
sub getRecurringDateFolderValues
{
    my $self = shift;
    my $extendedType = shift;
    my ($itemId, @folderIds) = @_;   
    
   return $self->getExtendedDateFolderValues("Recur", $itemId, @folderIds);
}

sub getRecurDateFolderValues
{
    my $self = shift;
    my $extendedType = shift;
    my ($itemId, @folderIds) = @_;   
    
   return $self->getExtendedDateFolderValues("Recur", $itemId, @folderIds);    
}

# This method can only take an array of dates.
# $itemId, $dateFolderId, 
sub getExtendedDateFolderValues
{
    my $self = shift;
    my $extendedType = shift;
    my ($itemId, @folderIds) = @_;   
    
    my $folders = "";
    my $sep = "";
    for my $folderId (@folderIds) 
    {
        $folders .= $sep . $folderId . ";$extendedType";
        $sep = ",";
    }
    
    my $cmd = "GetFolderValues,$itemId\r$folders";
    print ("DEBUG getFolderValues: $cmd\n") if ($DEBUG);
        
    my $items = request($self,$cmd); 
    my @itemList = split(/,/, $items, ($#folderIds+1));
    print ("DEBUG getFolderValues RETURN: [$items] -> [@itemList] = $#itemList\n") if ($DEBUG);
    
    return @itemList;
}


sub processStartDate
{
    my $self = shift;
    my $eccoDate = shift;
    print "processStartDate: $eccoDate\n" if ($DEBUG);
    
    croak("processStartDate: No value defined for eccoDate") unless (defined($eccoDate));  
    
    my $rawStartDate = undef;
    if ($eccoDate =~ /^(\d{4})(\d{2})(\d{2})$/) # date only
    {
       $rawStartDate = $eccoDate;
    } else 
    {
        my @dateValues = $self->getDateValues($eccoDate);
        
        if (! defined($dateValues[3])) 
        {
            $rawStartDate = $eccoDate;
        } else
        {
            $rawStartDate = $self->formatDate(@dateValues);
            $rawStartDate = Date::ICal->new(ical => $rawStartDate, offset => "-0800")->ical;
        }
    }     
    return $rawStartDate;
}

sub processEndDate
{
    my $self = shift;
    my $eccoDate = shift;
    my @startDateValues = @_;
    
    print "processEndDate: $eccoDate" . scalar(@startDateValues) . "\n" if ($DEBUG);
    
    croak("processEndDate: No value defined for eccoDate") unless (defined($eccoDate));  
    
    croak("processEndDate: No start date values defined") unless ($#startDateValues > -1);
    
    my $rawEndDate = "";
    if ($eccoDate =~ /(\d{4})(\d{2})(\d{2})/) # date only
    {
       $rawEndDate = $eccoDate;
    } elsif ($eccoDate =~ /(\d{2})(\d{2})/) # time only 
    {
        my @endDateValues = @startDateValues;
        $endDateValues[3] = $1;
        $endDateValues[4] = $2;
        my $formattedDate = $self->formatDate(@endDateValues);
        print "processEndDate: endDateValues = " . scalar(@endDateValues) . "\n" if ($DEBUG);
        $rawEndDate = Date::ICal->new(ical => $formattedDate, offset => "-0800")->ical;
        # print "processEndDate: formattedDate = $formattedDate, rawEndDate = $rawEndDate\n";
    } else # date + time
    {
        my @endDateValues = $self->getDateValues($eccoDate);
        print "processEndDate: endDateValues = " . scalar(@endDateValues) . "\n" if ($DEBUG);
        $rawEndDate = $self->formatDate(@endDateValues);
        $rawEndDate = Date::ICal->new(ical => $rawEndDate, offset => "-0800")->ical;
    }      
    
    return $rawEndDate;
}

sub getDateValues
{
    my ($self, $dateString) = @_; 
    
    print "getDateValues: $dateString\n" if ($DEBUG);
    
    croak("getDateValues: No value defined for dateString") unless (defined($dateString));
    
    my @datevalues = ($dateString =~ /(\d{4})(\d{2})(\d{2})(\d{0,2})(\d{0,2})/);
    return @datevalues;
}

sub formatDate
{
    if ($#_ != 5) 
    {
        die "formatDate: invalid number of parameters: " . join(' ', @_);
    }
    
    my ($self, $yyyy, $mm, $dd, $hours, $minutes) = @_;

    print "formatDate: $yyyy, $mm, $dd, $hours, $minutes\n" if ($DEBUG);
    
    my $formattedDate = "$yyyy$mm$dd";
    if (defined($hours)) {
      $formattedDate .= "T$hours$minutes";
    }
    
    return $formattedDate;    
}


# Recurring events are broken in Ecco Pro.  They set themselves with an end
# date of 12/31/1999, and so the only way you can get them working is by 
# setting a custom field.
sub parseRecurring
{
    my $self = shift;
    my $recurLine = shift;
    my ($notes, $type, $end, $interval, $pattern) = split(";", $recurLine);
          
    # <notes> is the "See All Recurring Notes" setting as a bool
    # print "notes = $notes,";
    $self->parseType($type);
    
    # <end> is the recurrence end date as YYYYMMDD
    # print " end = $end,\n";
    
    $self->parseInterval($interval);
 
    $self->parsePattern($pattern);
}

# type
# 1 - daily
# 2 - weekly
# 3 - monthly date
# 4 - monthly weekday
# 5 - monthly work day
# 6 - yearly                
# print " type = $type, "
sub parseType
{
    
}


#<interval> is the number of base periods (e.g., months for types 3, 4, and 5) between alarm sets, 
# print " interval = $interval,\n";
sub parseInterval
{
    
}


# <pattern> is a bit vector specifying the recurrence pattern encoded as a sequence of up to 12 hexadecimal digits 
# (i.e., up to 6 bytes as 2-digit hex numbers), the pattern is not used for types 1 and 6 and has the
# following meaning for the other types
# weekly bits 0 through 6 for the days SUN though SAT
# month/date bits 0 through 30 for the 1st through the 31st day of the month
# month/weekday bits 0 through 6 of byte 0 for 1st SUN though 1st SAT, 1. bits 0 through 6 of byte 1 for 2nd SUN through 2nd SAT,
# bits 0 through 6 of byte 4 for 5th SUN though 5th SAT,
# bits 0 through 6 of byte 5 for last SUN though last SAT,
# month/workday bit 0 set if last working day and clear if first working day

# print " pattern = $pattern ";
sub parsePattern
{
    
}

sub DESTROY {
    my $self = shift;
    my $conn = $self->{CONNECTION};
    if (defined($conn)) 
    {        
        $conn->Disconnect;    
    }
}

1;
__END__
# Below is stub documentation for your module. You'd better edit it!

=head1 NAME

Ecco - Perl extension for blah blah blah

=head1 SYNOPSIS

  use Ecco;
  blah blah blah

=head1 DESCRIPTION

Stub documentation for Ecco, created by h2xs. It looks like the
author of the extension was negligent enough to leave the stub
unedited.

Blah blah blah.

=head2 EXPORT

None by default.

=head1 SEE ALSO

Mention other useful documentation such as the documentation of
related modules or operating system documentation (such as man pages
in UNIX), or any relevant external documentation such as RFCs or
standards.

If you have a mailing list set up for your module, mention it here.

If you have a web site set up for your module, mention it here.

=head1 AUTHOR

Will Sargent, <lt>will_sargent@yahoo.com<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2005 by Will Sargent

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself, either Perl version 5.8.4 or,
at your option, any later version of Perl 5 you may have available.


=cut

