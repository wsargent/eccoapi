package Win32::Ecco::Import::OML;

use strict;

use Win32::Ecco;
use Carp;
use XML::DOM;
use Date::Parse;
use Date::Format;

use Exporter;
use vars qw($VERSION @ISA @EXPORT);

$VERSION= 1.00;

@ISA = qw(Exporter);

# Symbols to autoexport
@EXPORT =  qw( importOML );  

my $startFolderName = "Start Dates";
my $dueFolderName = "Due Dates";
my $finishFolderName = "Done";

my $DEBUG = 0;

sub new 
{
    my $self = {};
    my $class = shift;
    bless($self, $class);
    $self->{ECCO} = shift;
    return $self;
}


sub importOML
{
    my $self = shift;
    my $folderName = shift;
    my $filename = shift;
    
    print "importOML: $self, $folderName, $filename\n" if ($DEBUG);
    
    my @parents;
    
    my $ecco = $self->{ECCO};
    my $folderId = $ecco->getFolderId($folderName);
    die "Cannot find folder: $folderName" unless defined($folderId);
 
    my $parser = new XML::DOM::Parser;
    my $doc = $parser->parsefile($filename);
    $self->parseDocument($doc, $folderId);
}


sub parseOutline
{
    my $self = shift;
    my $outline = shift;
    my $folderId = shift;
    
    print "parseOutline: $self, $outline, $folderId\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
    
    my @parents = @_;
    
    my $text = $outline->getAttribute("text");
    
    # If there are any internal quotes, we have to double them so that they
    # get passed through appropriately.
    $text =~ s/\"/\"\"/g;
    
    # Quote it so we can deal with commas.
    $text = "\"" . $text . "\"";
    
    my $itemId = $ecco->createItem($folderId, $text);
    # For some wierd reason, create item doesn't set it to the folder
    # automatically.
    $ecco->setFolderValues($itemId, $folderId, 1);
    
    my $created = $outline->getAttribute("created");
    
    my $parentId = pop @parents;
    # insert the node if it has a parent.
    if (defined($parentId))
    {
        $ecco->insertItem($parentId, "d", $itemId);    
    }
    
    my $nodes = $outline->getChildNodes();
    for (my $i = 0; $i < $nodes->getLength(); $i++)
    {
         my $node = $nodes->item($i);
         if ($node->getNodeType() == ELEMENT_NODE)
         {
             my $nodeName = $node->getNodeName();
             
             if ($nodeName eq "data")
             {
                $self->parseData($itemId, $folderId, $node);    
             }
             
             if ($nodeName eq "item")
             {
                $self->parseItem($itemId, $folderId, $node);    
             }
             
             if ($nodeName eq "outline")
             {
                 # Add the current item as the parent.
                 push (@parents, $itemId);
                 $self->parseOutline($node, $folderId, @parents);    
             }             
         }
    }
}

sub parseDocument($$$)
{
    my $self = shift;
    my $doc = shift;
    my $folderId = shift;
    
    die "No folderId specified!" unless defined($folderId);

    print "parseDocument: $self, $doc, '$folderId'\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
    
    my $oml = $doc->getDocumentElement();
    my $nodes = $oml->getChildNodes();
    my @parents = undef;
    
    for (my $i = 0; $i < $nodes->getLength(); $i++)
    {
         my $node = $nodes->item($i);
         if ($node->getNodeType() == ELEMENT_NODE)
         {
             my $nodeName = $node->getNodeName();
             if ($nodeName eq "body")
             {
                 $self->parseBody($node, $folderId, @parents);    
             }
         }
    }
}

sub parseBody
{
    my $self = shift;
    my $body = shift;
    my $folderId = shift;
    my @parents = @_;
    
    print "parseBody: $self, $body, '$folderId'\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
    
    my $nodes = $body->getChildNodes();
    for (my $i = 0; $i < $nodes->getLength(); $i++)
    {
         my $node = $nodes->item($i);
         if ($node->getNodeType() == ELEMENT_NODE)
         {
             my $nodeName = $node->getNodeName();
             if ($nodeName eq "outline")
             {
                 $self->parseOutline($node, $folderId, @parents);    
             }
         }
    }
}

sub parseItem($$$)
{
    my $self = shift;
    my $itemId = shift;
    my $folderId = shift;
    my $item = shift;
    
    print "parseItem: $self, $itemId, $folderId, $item\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
    
    my $propertyName = $item->getAttribute("name");
    my $propertyValue = $self->getTextValue($item);
    
    # Add the dates from the shadow item.  We don't really care about the
    # created tag, we want the Ecco created tag anyway...
    # Use OML date format.
    my $startFolderId = $ecco->getFolderId($startFolderName);
    my $dueFolderId = $ecco->getFolderId($dueFolderName);
    my $finishFolderId = $ecco->getFolderId($finishFolderName);
    
    my $dateFormat = "%a, %e %b %Y %X %Z";
    if ($propertyName eq $startFolderName)
    {
        my $eccoTime = time2str($dateFormat, str2time($propertyValue));  
        $ecco->setFolderValues($itemId, $startFolderId, $eccoTime);
    } elsif ($propertyName eq $finishFolderName)
    {
        my $eccoTime = time2str($dateFormat, str2time($propertyValue));  
        $ecco->setFolderValues($itemId, $finishFolderId, $eccoTime);
    } elsif ($propertyName eq $dueFolderName)
    {
        my $eccoTime = time2str($dateFormat, str2time($propertyValue));  
        $ecco->setFolderValues($itemId, $dueFolderId, $eccoTime);
    } else
    {
        # See if we have a folder by this name.  This is a last ditch effort
        # to match a shadow tag with an Ecco folder.
        my $tagFolderId = $ecco->getFolderId($propertyName);
        if (defined($tagFolderId))
        {
            # it pretty much has to be a boolean flag, so set it to true.
            $ecco->setFolderValues($itemId, $tagFolderId, "1");
        }
    }
}

sub parseData
{
    my $self = shift;
    my $parentId = shift;
    my $folderId = shift;
    my $data = shift;
    
    print "parseData: $self, $parentId, $folderId, $data\n" if ($DEBUG);
    
    my $ecco = $self->{ECCO};
    
    my $text = $self->getTextValue($data);
    
    my @lines = split(/\n/, $text);
    for my $line (@lines)
    {
       # chomp $line;
       # remove any empty lines.
       $line =~ s/^\s+$//;
        
       # print "line = $line\n";
       if ($line ne "")
       {
           my $itemId = $ecco->createItem($folderId, $line);      
           $ecco->insertItem($parentId, "d", $itemId) if (defined($itemId));
       }
    }
}

sub getTextValue($)
{
    my $self = shift;
    my $element = shift;
    
    print "parseData: $self, $element\n" if ($DEBUG);
    
    my $nodes = $element->getChildNodes();
        
    # Make a buffer of text and CDATA nodes.
    my $b;
    for (my $i = 0; $i < $nodes->getLength(); $i++)
    {
        my $node = $nodes->item($i);
        if ($node->getNodeType() == TEXT_NODE)
        {
            $b .= $node->getNodeValue();
        }

        if ($node->getNodeType() == CDATA_SECTION_NODE)
        {
           $b .= $node->getNodeValue();
        }
    }

    return $b;
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

