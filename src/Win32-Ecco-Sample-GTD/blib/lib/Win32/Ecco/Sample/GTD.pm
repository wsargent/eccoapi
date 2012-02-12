# This library contains useful methods for running through the 
# Getting Things Done library.

package Win32::Ecco::Sample::GTD;

use strict;
use Carp;

use Win32::Ecco::Import::OML;
use Win32::Ecco::Export::OML;

use File::Copy;

BEGIN {
  use Exporter;
  our ($VERSION, @ISA, @EXPORT, @EXPORT_OK, %EXPORT_TAGS);

  $VERSION=1.0;
  @ISA=qw(Exporter);
  @EXPORT=
        qw(&importDiary
        &importInbasket
        &importCommand
        &exportActions
        &exportProjects
        &exportSomeday
        &google);
  @EXPORT_OK=qw(%props);
}
our @EXPORT_OK;

my $DEBUG = 0;

sub new 
{
    my $self = {};
    my $class = shift;
    bless($self, $class);
    $self->{ECCO} = shift;    
    return $self;
}

sub importDiary
{
    my $self = shift;    
    my $omlFile = shift;
    my $shadowFile = shift;
    
    print "importDiary: $self, $omlFile, $shadowFile\n" if ($DEBUG);
    my $ecco = $self->{ECCO};
    
    # `$HOME/ecco/import.pl "Diary" "$HOME/$filename.xml"`;
    my $omlImportVar = Win32::Ecco::Import::OML->new($ecco);
    $omlImportVar->importOML("Diary", $omlFile);   
    deleteImportedItems($shadowFile);
}

sub importInbasket
{        
    # `$HOME/ecco/import.pl "Diary" "$HOME/$filename.xml"`;
    my $self = shift;    
    my $omlFile = shift;
    my $shadowFile = shift;
    
    print "importInbasket: $self, $omlFile, $shadowFile\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
    
    my $omlImportVar = Win32::Ecco::Import::OML->new($ecco);
    
    $omlImportVar->importOML("Inbasket", $omlFile);
    
    deleteImportedItems($shadowFile);
}

# Assumes all the folders take check boxes.
# importCommand($text, @folderNames);
sub importCommand
{
    my $self = shift;
    my $text = shift;
    my $folderName = shift;
    my @folderNames = @_;
    
    print "importCommand: $self, $text, $folderName\n" if ($DEBUG);
    
    croak("self is not defined") unless (defined($self));
    croak("text is not defined") unless (defined($text));
    croak("folderName is not defined") unless (defined($folderName));
    
    my $ecco = $self->{ECCO};
    
    my $folderId = $ecco->getFolderId($folderName);
    croak("importCommand: no value for folderId using $folderName") unless defined($folderId);
    
    my $itemId = $ecco->createItem($folderId, $text);
    croak("importCommand: no value for itemId using $folderName") unless defined($itemId);
        
    # Hmm, seems like it needs to be set again.
    $ecco->setFolderValues($itemId, $folderId, 1);
    
    for my $folderName (@folderNames)
    {
        $folderId = $ecco->getFolderId($folderName);
        if (defined($folderId)) 
        {
            $ecco->setFolderValues($itemId, $folderId, 1);
        }
    }
}

sub exportActions
{
    my $self = shift;    
    my $shadowFileName = shift,
    my $omlFileName = shift;
    
    print "exportActions: $self, $shadowFileName, $omlFileName\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
                    
    # Get the item properties ids I'm interested in...
    my $tseCreatedId = $ecco->getFolderId("TSeCreated");
    my $actionId = $ecco->getFolderId("Action");
    my $doneId = $ecco->getFolderId("Done");
    # my $notId = $ecco->getFolderId("Not");   
    my $waitingId = $ecco->getFolderId("Waiting For");

    # Set up a filter.
    my %filter = ( $actionId => 1, $doneId => undef(), $waitingId => undef() );
    
    my $viewName = "Actions";
    my %viewHash = $ecco->getViewsHash();
    my $viewId = $viewHash{$viewName};
    
    my %metadata = (
         viewId => $viewId, viewName => $viewName
    );
    
    local *FH;
    open(FH, "> $omlFileName") or die "Can't open $omlFileName: $!";
    my $filehandle = *FH;
    $self->printOML(\%metadata, $viewId, \%filter, $filehandle);
    close(FH) or die "Can't close $omlFileName: $!";
}

sub exportProjects
{
    # Create a new ECCO connection
    my $self = shift;
    my $shadowFileName = shift,
    my $omlFileName = shift;
    
    print "exportProjects: $self, $shadowFileName, $omlFileName\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
    
    my %filter = ();
    my $viewName = "Projects";
    my %viewHash = $ecco->getViewsHash();
    my $viewId = $viewHash{$viewName};
    
    my %metadata = (
         viewId => $viewId, viewName => $viewName
    );
 	
    local *FH;
    open(FH, "> $omlFileName") or die "Can't open $omlFileName: $!";
    my $filehandle = *FH;
    $self->printOML(\%metadata, $viewId, \%filter, $filehandle);
    close(FH) or die "Can't close $omlFileName: $!";
}


sub exportSomeday
{
    my $self = shift;
    my $shadowFileName = shift,
    my $omlFileName = shift;
    
    print "exportSomeday: $self, $shadowFileName, $omlFileName\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
    
    # Get the item properties ids I'm interested in...
    my $tseCreatedId = $ecco->getFolderId("TSeCreated");
    my $actionId = $ecco->getFolderId("Action");
    my $doneId = $ecco->getFolderId("Done");
    # my $notId = $ecco->getFolderId("Not");            

    # Not actionable, not done
    my %filter = ( $actionId => undef(), $doneId => undef() );
    
    # Get the view.
    my $viewName = "Someday";
    my %viewHash = $ecco->getViewsHash();
    my $viewId = $viewHash{$viewName};
    
    my %metadata = (
         viewId => $viewId, viewName => $viewName
    );
    	
    local *FH;
    open(FH, "> $omlFileName") or die "Can't open $omlFileName: $!";
    my $filehandle = *FH;
    $self->printOML(\%metadata, $viewId, \%filter, $filehandle);
    close(FH) or die "Can't close $omlFileName: $!";
}

sub printOML
{
    my $self = shift;
    my $metadataRef = shift;
    my $viewId = shift;
    my $filterRef = shift;
    my $filehandle = shift;
    
    print "printOML: $self, $metadataRef, $viewId, $filterRef, $filehandle\n" if ($DEBUG);
    
    my %metadata = %{$metadataRef};
    my %filter = %{$filterRef};
    
    my $ecco = $self->{ECCO};
    
    my $oml = Win32::Ecco::Export::OML->new($ecco, $filehandle);
    
    $oml->printHeader(%metadata);
    my @folders = $ecco->getViewTLIs($viewId);
    for my $i (0 .. $#folders) 
    {
         my @tli = @{$folders[$i]};
         my $folderId = shift @tli;
         $oml->printFolder($folderId, \@tli, \%filter);                                 
    }   
    $oml->printFooter();
}

sub deleteImportedItems 
{
    my ($oldfile) = @_;

    print "deleteImportedItems: $oldfile\n" if ($DEBUG);
    
    # Move the items over as a group (probably safer that way).
    my $newfile = "$oldfile.out";
    open(FILE,"<$oldfile") or die "Can't find file: $oldfile $!\n";
    open(OUT, ">$newfile") or die "Can't create $newfile: $!\n";
    while (<FILE>)
    {
        # marking the items as deleted is safer than deleting the file,
        # and lets shadow know how to sync when both the handheld and
        # the desktop are dirty.
	s/deleted="0"/deleted="1"/g;
	print OUT $_;
    }
    close OUT;

    # overwrite at close
    move($newfile, $oldfile);
}


sub google
{
    my $self = shift;
    my $itemId = shift;
    
    print "printOML: $self, $itemId\n" if ($DEBUG);
    
    # Create a new ECCO connection
    my $ecco = $self->{ECCO};
    # print "itemId = $itemId\n";

    my $text = $ecco->getItemText($itemId);
    $text = $ecco->dequote($text);
    $text = url_encode($text);
    
    # my $program = "rundll32 url.dll,FileProtocolHandler";
    my $program = "c:/program files/Mozilla Firefox/firefox.exe";
    my $cmd = "\"$program\" http://www.google.com/search?q=$text";
    `$cmd`;
}


sub url_encode {
    my $text = shift;
    $text =~ s/([^a-z0-9_.!~*'(  ) -])/sprintf "%%%02X", ord($1)/ei;
    $text =~ tr/ /+/;
    return $text;
}



1;
