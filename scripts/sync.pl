#!/usr/perl/bin/perl

# !!!WARNING!!!!
# Note that this script has some issues which prevent it from being used.
# Items that are deleted can popup again with a new ID, and I don't know why.
# !!!WARNING!!!!

# Looks at the shadow file, and tells us that anything that's been checked
# (which is a shadow operation) should be marked as done on the Ecco side.
#
# This prevents us from tediously checking off things done in both Shadow 
# and Ecco.

use strict;
use warnings;

die "This script is buggy.  Do not use.";

use FindBin qw($Bin);
use lib "$Bin";
use Win32::Ecco; 

use XML::DOM;

use Date::Parse;
use Date::Format;

my $HOME = $ENV{HOME};
my $filename = $ARGV[0];
my $context = $ARGV[1];

my $ecco = Win32::Ecco->new();

my $finishFolderName = "Done";
my $finishFolderId = $ecco->getFolderId($finishFolderName);

my $pilotRecordName = "$context Record ID";
my $pilotFolderId = $ecco->getFolderId($pilotRecordName);

my $PALM_EPOCH_OFFSET = 2082844800;

# main
{
    my $mod = (stat "$filename")[9];
 
    # this prints result in usual date format
    # like -> Tue Jun 12 23:50:10 2001
    
    my $dateMod = scalar localtime $mod;
    print "dateMod = $dateMod\n";
    
    my $parser = new XML::DOM::Parser;
    my $doc = $parser->parsefile($filename);
    parseDocument($doc);
}

sub parseDocument
{
    my $doc = shift;
    
    my $shadowPlanFile = $doc->getDocumentElement();
    my $uniqueTime = $shadowPlanFile->getAttribute("uniqueTime");
    
    # print "uniqueTime = $uniqueTime\n";
    
    my $nodes = $shadowPlanFile->getChildNodes();
    for (my $i = 0; $i < $nodes->getLength(); $i++)
    {
         my $node = $nodes->item($i);
         if ($node->getNodeType() == ELEMENT_NODE)
         {
             my $nodeName = $node->getNodeName();
             if ($nodeName eq "item")
             {
                 parseItem($node);    
             }
         }
    }
}

sub parseItem
{
    my $item = shift;
    my $deleted = $item->getAttribute("deleted");
    
    # If it's been deleted, we don't want to update it.
    return if ($deleted);
   
    my $itemId = $item->getAttribute("localID");
    my $uniqueId = $item->getAttribute("uniqueID");
    
    # If it's zero, we know this is after the export...
    if ($uniqueId ne "0")
    {
        $ecco->setFolderValues($itemId, $pilotFolderId, $uniqueId);
    }
    
    my $dirtyContent = $item->getAttribute("dirtyContent");
    my $checked = $item->getAttribute("checked");
    if ($checked eq "yes" and $dirtyContent ne "1")
    {
       my $text = getChildText($item, "title");
       my $shadowFinishTime = getChildText($item, "hhFinishTime");
       my $tse = ($shadowFinishTime - $PALM_EPOCH_OFFSET);

       # Format it as a GMT oriented epoch.
       my $eccoTime = time2str("%Y%m%d%H%M", $tse, "GMT");
       print "localID = $itemId, checked = $checked, finishTime = $eccoTime, text = \"$text\"\n";
       if (defined($eccoTime))
       {
           $ecco->setFolderValues($itemId, $finishFolderId, $eccoTime);
       }
    }

    my $nodes = $item->getChildNodes();
    for (my $i = 0; $i < $nodes->getLength(); $i++)
    {
         my $node = $nodes->item($i);
         if ($node->getNodeType() == ELEMENT_NODE)
         {
             my $nodeName = $node->getNodeName();
             if ($nodeName eq "item")
             {
                 parseItem($node);    
             }
         }
    }
}

sub getChildText()
{
    my $element = shift;
    my $child = shift;

    my $nodes = $element->getChildNodes();
    for (my $i = 0; $i < $nodes->getLength(); $i++)
    {
         my $node = $nodes->item($i);
         if ($node->getNodeType() == ELEMENT_NODE)
         {
             my $nodeName = $node->getNodeName();
             if ($nodeName eq $child)
             {
                 return getTextValue($node);  
             }
         }
    }
    return;
}

sub getTextValue()
{
    my $element = shift;
    
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
