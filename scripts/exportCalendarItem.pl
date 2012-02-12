#!/usr/perl/bin/perl -w

=head2 Discussion

This script shows how to export data from Ecco Pro to an iCalendar file

Will Sargent (will_sargent@yahoo.com)

=cut

use strict;
use warnings;

use Win32::Ecco;
use Win32::Ecco::Export::ICal;
use Data::ICal;

# Makes subprocesses not show console windows.  This only works in ActivePerl...
#BEGIN {
#    Win32::SetChildShowWindow(0) if (defined &Win32::SetChildShowWindow)
#};

# main
{
    my $itemId = shift @ARGV;
    
    print "itemId = $itemId\n";
    
    my $ecco = Win32::Ecco->new();
    
    # Get the item properties ids I'm interested in...
    my $folderName = "Appointments";
    my $folderId = $ecco->getFolderId($folderName);

    die "folderId not found in folder $folderName" unless (defined($folderId));
    
    # Create a calendar, then print out the item in that folder.
    my $calendar = Data::ICal->new();
    my $exporter = Win32::Ecco::Export::ICal->new($ecco);   
    $exporter->exportCalendarItem($folderId, $itemId, $calendar);
    
    my $HOME = $ENV{HOME};
    open(FILE, "> $HOME/export.ics") or die "Can't write to file: $!";
    print FILE $calendar->as_string;    
    close FILE;

    # Display the calendar to the user so I have some idea of what's happening.
    print $calendar->as_string;    
    sleep 2;    
}

