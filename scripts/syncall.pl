#!/usr/perl/bin/perl -w

=head2 Discussion

This is a sample script that shows how Ecco Pro can be integrated into other
systems.  The GTD module imports two files from Shadow to Ecco, and exports
two files from Ecco to Shadow.

The GTD module in turn makes use of Ecco::Import::OML and Ecco::Export::OML. 

You should set SHADOWCONV_USERNAME and SHADOWCONV_HOME as environment variables
if you want to change the settings, or just edit this file to your preference.

Will Sargent (will_sargent@yahoo.com)

=cut

use strict;
use warnings;

use Win32::Ecco;
use Win32::Ecco::Sample::GTD;

# Makes subprocesses not show console windows.  This only works in ActivePerl...
#BEGIN {
#    Win32::SetChildShowWindow(0) if (defined &Win32::SetChildShowWindow)
#};

# Get the palm username from the environment variable
my $username = $ENV{'SHADOWCONV_USERNAME'};
if (! defined($username)) 
{
    die "You must specify your Palm username!"; 
}

# Get the shadowconverter home from the environment variable.
my $shadowConvHome = $ENV{'SHADOWCONV_HOME'};
if (! defined($shadowConvHome)) 
{
    die "You must specify your shadowConverter directory!";  
} 

my $shadowVersion = "ShadowPlan400";

my $omlDir = "$shadowConvHome";
my $propertiesDir = "$shadowConvHome/ecco";
my $converterFile = "$shadowConvHome/lib/converter.jar";
my $shadowDir = "c:\\program files\\Palm\\$username\\$shadowVersion";
my $eccoFile = "ecco.eco";

# main
{
    my $ecco = Win32::Ecco->new;
    
    # Check that we've got the right file open...
    # my $filename = $ecco->getFileName($ecco->getCurrentFile());
    # die "Don't have the right file: $filename\n" unless ($filename =~ /$eccoFile/);
        
    my $gtd = Win32::Ecco::Sample::GTD->new($ecco);
    
    #### Imports
    
    executeConverter("diary.properties");  
    $gtd->importDiary(getOmlFile("Diary"), getShadowFile("Diary"));
    
    executeConverter("inbasket.properties");
    $gtd->importInbasket(getOmlFile("Inbasket"), getShadowFile("Inbasket"));
        
    #### Exports
    
    $gtd->exportActions(getShadowFile("Actions"), getOmlFile("Actions"));
    executeConverter("actions.properties");
    
    $gtd->exportSomeday(getShadowFile("Someday"), getOmlFile("Someday"));          
    executeConverter("someday.properties");     
}

# Runs the java program which converts Shadow XML to and from OML.
sub executeConverter
{
    my $filename = getPropertiesFile(shift);
    my $cmd = "javaw -jar $converterFile $filename";
    `$cmd`;       
}

sub getPropertiesFile
{
    my $propertiesFile = shift;
    return "$propertiesDir/$propertiesFile";   
}

sub getShadowFile 
{
   my $filename = shift;
   return "$shadowDir\\ShadP-$filename.XML"	
}

sub getOmlFile
{
   my $filename = shift;
   return "$omlDir/$filename.xml";
}
