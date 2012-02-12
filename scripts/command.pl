#!/usr/perl/bin/perl 

# Imports an item into Ecco from a command line.

use strict;

use FindBin qw($Bin $Script);
use lib "$Bin";

use Win32;
use Win32::Process;
use Win32::Ecco; 
use Win32::Ecco::Sample::GTD;

# Run a detached command in the background, in a different process.
# We do this because otherwise the Yahoo Widget Engine is unbearably slow.

# Unfortunately it requires that I know where the Perl executable is, and I 
# really can't remember the magic incantation for that...
my $perl = "c:\\usr\\perl\\bin\\perl.exe";

{
    my $firstArg = $ARGV[0];
    
    if ($firstArg eq "-execute") 
    {
        shift @ARGV;
        my $text = join(' ', @ARGV);        
        importInbasket($text);
        
    } else 
    {
        # Get the path to the currently running executable.
        my $myExe = "$Bin/$Script";
        $myExe =~ s#/#\\#g;
        my $args = join(' ', @ARGV);  
        
        my $cmd = "$perl $myExe -execute \"$args\"";
        
        print "$cmd \n";
        
        Win32::Process::Create($Win32::Process::Create::ProcessObj,
            $perl,                   # Whereabouts of Perl
            $cmd,  #
            0,                                  # Don't inherit.
            DETACHED_PROCESS,                   #
            ".") or                             # current dir.
        die print_error();
    }
}
    
# Text is always the last element.
sub importInbasket
{
    my $folderName = "Inbasket";
    my $text = shift;
    
    print "importInbasket $text\n";
    
    my $ecco = Win32::Ecco->new();
    my $gtd = Win32::Ecco::Sample::GTD->new($ecco);
    
    $gtd->importCommand($text, $folderName);
}

sub print_error() {
    return Win32::FormatMessage( Win32::GetLastError() );
}
