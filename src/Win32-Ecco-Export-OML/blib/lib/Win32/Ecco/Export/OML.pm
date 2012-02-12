################################################################################
# OML Export
################################################################################
#
# Writes out all the items in the folder in OML format.  
#
#
#
#

package Win32::Ecco::Export::OML;

use strict;
use warnings;
use Carp;

use Win32::Ecco;

use Date::Parse;
use Date::Format;

use Exporter;
use vars qw($VERSION @ISA @EXPORT);

$VERSION=1.00;
@ISA=qw(Exporter);
@EXPORT=
    qw(&printItem 
    &escapeHTML 
    &printItemStart 
    &printItemEnd 
    &printFolder
    &printHeader
    &printFooter
    &printIndent);

my %FOLDER_TYPES = (
	checkmark => 1,
	date => 2,
	number => 3,
	text => 4,
	popup => 5,
);

my $DEBUG = 0;

sub new {
    my $self = {};
    my $class = shift;
    bless($self, $class);
    $self->{ECCO} = shift; 
    $self->{FILE} = shift;
    return $self;
}

# This doesn't do a perfect job of filtering (it should really do a depth
# first traversal for valid items and work up to the root, then back down again)
# but it's possible to set up filters in Shadow Plan that will deal with the
# edge cases so I really don't care that much.
sub printItem
{
    my $self = shift;
    my $indent = shift;
    my $itemId = shift;
    my $filterRef = shift;
    
    print "printItem: $self, $indent, $itemId, $filterRef\n" if ($DEBUG);
    
    if (! defined($itemId)) 
    {
        die "itemId not defined!";    
    }
    
    my $ecco = $self->{ECCO};
    my @children = $ecco->getItemSubs(1, $itemId);     
    if ($#children > -1) # if there any children...
    {
        # We always print out parents, even if they don't have children
        # that match...
        $self->printItemStart($indent, $itemId);
                      
        for my $i ( 0 .. $#children ) 
        {
            my $aref = $children[$i];
            if (defined($aref)) 
            {
                my $n = @$aref - 1;
                for my $j ( 0 .. $n ) 
                {                              
                    my $subItemId = $children[$i][$j];
                    $self->printItem($indent + 1, $subItemId, $filterRef);
                }
            }
        }
        $self->printItemEnd($indent);
    } else
    {
        # if an item has no children, then see if we can filter it.
        if ($ecco->filterItem($itemId, %{$filterRef}))
        {
            $self->printItemStart($indent, $itemId);
            $self->printItemEnd($indent);
        }
    }
}

sub escapeHTML
{
    my $self = shift;
    my $text = shift;
    
    print "printItem: $self, $text\n" if ($DEBUG);
    
    # Convert the ampersand character into its character entity
    $text =~ s/&/&amp;/g;       
    
    # replace all quotes left with the character entity
    $text =~ s/\"/&quot;/g;
    
    # Change any < and > characters to their equivalent entities.
    $text =~ s/</&lt;/g;
    $text =~ s/>/&gt;/g;
    
    # Convert any high valued characters into their numeric character entities
    $text =~ s/([^\x20-\x7F])/'&#' . ord($1) . ';'/gse;
    
    return $text;
}

sub printItemStart
{
    my $self = shift;
    my $indent = shift;
    my $itemId = shift;
    
    print "printItemStart: $self, $indent, $itemId\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
    my $text = $self->escapeHTML($ecco->dequote($ecco->getItemText($itemId)));      
    
    my $FILE = $self->{FILE};
    $self->printIndent($indent);
    print $FILE "<outline text=\"$text\">\n";
    $self->printIndent($indent);
    print $FILE "\t<item name=\"itemId\">$itemId</item>\n";

    my @itemFolders = $ecco->getItemFolders($itemId);
    my @folderValues = $ecco->getFolderValues($itemId, @itemFolders);
   
    for my $i (0 .. $#itemFolders)
    {
        my $folderId = $itemFolders[$i];
        my $folderValue = $folderValues[$i];
	$folderValue = $self->escapeHTML($ecco->dequote($folderValue));   
	
	# Find out what the folder type is, so we can convert it from 
	# Ecco date format to OML date format if necessary.
        my $folderType = $ecco->getFolderType($folderId);
        die "Undefined folder type in folderId $folderId" unless defined($folderType);
        
	if ($folderType eq $FOLDER_TYPES{'date'}) {
	     # "EE, d MMM yyyy HH:mm:ss z" in Java format...
	     $folderValue = time2str("%a, %e %b %Y %X %Z", str2time($folderValue));
	}
	
        my $folderName = $self->escapeHTML($ecco->dequote($ecco->getFolderName($folderId)));
        $self->printIndent($indent);
        print $FILE "\t<item name=\"$folderName\">$folderValue</item>\n";
    }  
}

sub printItemEnd
{
    my $self = shift;
    my $indent = shift;
    
    print "printItemStart: $self, $indent\n" if ($DEBUG);
    
    $self->printIndent($indent);
    
    my $FILE = $self->{FILE};
    print $FILE "</outline>\n";
}

sub printFolder
{
    my $self = shift;
    my $folderId = shift;
    
    my $tliRef = shift;
    my $filterRef = shift;
    
    print "printItemStart: $self, $folderId, $tliRef, $filterRef\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
    my $folderName = $ecco->dequote($ecco->getFolderName($folderId));
    my $folderType = $ecco->getFolderType($folderId);
    
    my $FILE = $self->{FILE};
    print $FILE "\t<outline text=\"$folderName\">\n";
    print $FILE "\t\t<item name=\"folderId\">$folderId</item>\n";
    print $FILE "\t\t<item name=\"folderType\">$folderType</item>\n";
    
    # If printing out a folder, mark it with the name so we know to tag
    # the folder correctly (note that we only do this for checkmarks)
    if ($folderType eq "1")
    {
        print $FILE "\t\t<item name=\"$folderName\">1</item>\n";
    }
    
    for my $itemId (@{$tliRef})
    {
        $self->printItem(1, $itemId, $filterRef);
    }
    print $FILE "\t</outline>\n";
}

# FIXME should escape entities in metadata.
sub printHeader
{
    my $self = shift;
    my %metadata = @_;
    
    print "printHeader: $self, %metadata\n" if ($DEBUG);
    
    my $FILE = $self->{FILE};
    print $FILE <<"EOF";
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE oml SYSTEM "http://oml.sf.net/spec/oml.dtd">
EOF
    print $FILE "<oml>\n";
    print $FILE "<head>\n";
    for my $key (keys %metadata) 
    {
        if (defined($key)) 
        { 
            my $value = $metadata{$key};
            $value = "" if (! defined($value));
            print $FILE "\t<metadata name=\"$key\">$value</metadata>\n";
        }
    }
    print $FILE "</head>\n";
    print $FILE "<body>\n";
}

sub printFooter
{    
    my $self = shift;
    
    print "printFooter: $self\n" if ($DEBUG);
    
    my $FILE = $self->{FILE};
    print $FILE "</body>\n</oml>\n";    
}

sub printIndent
{
    my $self = shift;
    my $indent = shift;
    
    print "printFooter: $self, $indent\n" if ($DEBUG);
    
    my $FILE = $self->{FILE};
    print $FILE ("\t" x ($indent + 1));    
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
