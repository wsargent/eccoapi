#!/usr/perl/bin/perl

# Look up the item text in Google, then call a command from it.
# 
# You can get the item id from Ecco by defining a launch item this way:
# 
# c:\usr\perl\bin\perl.exe c:\home\wsargent\bin\google.pl "<item-id>"

use strict;

use FindBin qw($Bin);
use lib "$Bin";
use Win32::Ecco; 
use Win32::Ecco::Sample::GTD;
         
# Text is always the last element.
{
    my $itemId = pop @ARGV;
    
    my $ecco = Win32::Ecco->new();
    my $gtd = Win32::Ecco::Sample::GTD->new;
    
    $gtd->google($ecco, $itemId);
}
