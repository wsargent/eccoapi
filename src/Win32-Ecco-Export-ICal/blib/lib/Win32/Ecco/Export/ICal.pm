package Win32::Ecco::Export::ICal;

# Exports Ecco Pro data to iCalendar format.

# Needs documentation
# Needs WAY better date formatting
# Needs extended date information
# Needs better control (provide option to export only one item)

# will_sargent@yahoo.com

use strict;
use Carp;

use Win32::Ecco;

use Data::ICal;
use Date::ICal;
use Data::ICal::Entry::Event;

use File::Copy;

BEGIN {
  use Exporter;
  our ($VERSION, @ISA, @EXPORT, @EXPORT_OK, %EXPORT_TAGS);

  $VERSION=0.01;
  @ISA=qw(Exporter);
  @EXPORT=
        qw(
        &exportCalendar
        );
  @EXPORT_OK=qw(%props);
}
our @EXPORT_OK;

my $DEBUG = 0;

# Takes the class and a reference to Ecco.
sub new 
{
    my $self = {};
    my $class = shift;
    bless($self, $class);
    $self->{ECCO} = shift;    
    return $self;
}

# Exports all the calendar dates in the folder identified by calendarFolderId.
sub exportCalendar
{
    my $self = shift;
    my $calendarFolderId = shift;
    my $calendar = shift;
    
    print "exportCalendar: $calendarFolderId\n" if ($DEBUG);
    
    croak("exportCalendar: undefined calendarFolderId") unless defined($calendarFolderId);
    
    my $ecco = $self->{ECCO};
    
    my @items = $ecco->getFolderItems($calendarFolderId);
    
    
    for my $i (0 .. $#items) 
    {
        my $itemId = $items[$i];
        if (! defined($itemId)) 
        {
            die "No itemId defined for calendarFolderId $calendarFolderId\n";
        }
        
        $self->exportCalendarItem($calendarFolderId, $itemId, $calendar);
    }
}

sub exportCalendarItem
{
    my $self = shift;
    my $calendarFolderId = shift;
    my $itemId = shift;
    my $calendar = shift;

    my $ecco = $self->{ECCO};
    
    my $itemText = $ecco->getItemText($itemId);
    $itemText = $self->deq($itemText);
    
    my $description = $self->printDescription("", $itemId);
    
    # print "description = $description\n";
    
    my @itemFolders = $ecco->getItemFolders($itemId);
    my @folderValues = $ecco->getFolderValues($itemId, @itemFolders);
    
    for my $i (0 .. $#itemFolders)
    {
        my $itemFolderId = $itemFolders[$i];
        
        # Only print events for items that match up with our calendar date...
        if ($itemFolderId eq $calendarFolderId) 
        {
            my $folderValue = $folderValues[$i];
            $folderValue = $self->deq($folderValue); 
            
            print "exportCalendar: folderValue = $folderValue\n" if ($DEBUG);
            
            # If we find a dash in the appointments field, then we 
            # know it's a date range.
            my $itemStartDate;
            my $itemEndDate; 
            my @dateRange = split('-', $folderValue);
            if ($#dateRange > 0) 
            {
                print "exportCalendar: dateRange = $dateRange[0] to $dateRange[1]\n" if ($DEBUG);
                $itemStartDate = $ecco->processStartDate($dateRange[0]);
                my @dateValues = $ecco->getDateValues($dateRange[0]);
                $itemEndDate = $ecco->processEndDate($dateRange[1], @dateValues);
            } else 
            {
                $itemStartDate = $ecco->processStartDate($folderValue);
                $itemEndDate = $itemStartDate;
            }

            #my @recurringValues = $ecco->getRecurringDateFolderValues($itemId, $calendarFolderId);
            #if ($#recurringValues >= 0) {
            #    my $recurLine = shift @recurringValues;
            #    $self->parseRecurring($recurLine);
            #}
            
            print "itemText = $itemText, folderValue = $folderValue, itemStartDate = $itemStartDate, itemEndDate = $itemEndDate\n" if ($DEBUG);
            
            my $vevent = $self->createEvent($itemText, $itemStartDate, $itemEndDate, $description);
            $calendar->add_entry($vevent);
        }      
    }    
    
}

sub printDescription
{
    my $self = shift;
    my $text = shift;
    my $itemId = shift;
    
    print "printDescription: text = $text, itemId = $itemId\n" if ($DEBUG);
    
    if (! defined($itemId)) 
    {
        die "itemId not defined!";    
    }
    
    my $ecco = $self->{ECCO};
    
    my @children = $ecco->getItemSubs(1, $itemId);     
    if ($#children > -1) # if there any children...
    {
        for my $i ( 0 .. $#children ) 
        {
            my $aref = $children[$i];
            if (defined($aref)) 
            {
                my $n = @$aref - 1;
                for my $j ( 0 .. $n ) 
                {                              
                    my $subItemId = $children[$i][$j];
                    $text .= $self->printItemStart($subItemId);
                    # We always print out parents, even if they don't have children
                    # that match...
                    # $text .= printDescription($subItemId);
                }
            }
        }
    }
    
    return $text;
}


sub printItemStart()
{
    my $self = shift;
    my $itemId = shift;
    
    my $ecco = $self->{ECCO};
    my $text = $self->deq($ecco->getItemText($itemId));      

    return "$text ";
}


sub createEvent
{
    my ( $self, $itemText, $itemStartDate, $itemEndDate, $description) = @_;
    
    print "createEvent: $itemText, $itemStartDate, $itemEndDate, $description\n" if ($DEBUG);
    
    # Create a new ICal event.
    my $vevent = Data::ICal::Entry::Event->new();
    
    $vevent->add_properties(
            summary => $itemText,
            description => $description,
            dtstart => $itemStartDate,
            dtend => $itemEndDate
    );
    
    return $vevent;            
}

# Strips out the quotes from the text.
sub deq
{    
    my $self = shift;
    my $text = shift;

    $text =~ s/"(.*)"/$1/;
    
    # If there are quotes in the actual text, then they're replaced by two quotes.
    # So we reverse the process here.
    $text =~ s/""/"/g;

    return $text;    
}

1;
__END__
