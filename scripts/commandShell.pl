#!/usr/perl/bin/perl 

# Run a detached command in the background, in a different process.
# We do this because otherwise the Yahoo Widget Engine is unbearably slow.

use strict;
use Win32;
use Win32::Process;

my $command = "c:\\usr\\perl\\bin\\perl.exe";
my $path = "d:\\home\\wsargent\\work\\eccoapi\\scripts\\";
my $script = "$path\\command.pl";
my $args = join(' ', @ARGV);

Win32::Process::Create($Win32::Process::Create::ProcessObj,
    $command,            # Whereabouts of Perl
    "perl $script \"$args\"" ,  #
    0,                                  # Don't inherit.
    DETACHED_PROCESS,                   #
    ".") or                             # current dir.
die print_error();

sub print_error() {
    return Win32::FormatMessage( Win32::GetLastError() );
}
