use Win32::OLE;
use Win32::OLE::Const 'Microsoft Outlook';

my $outlook = Win32::OLE->new('Outlook.Application') or die "Error!\n";

my $namespace = $outlook->GetNamespace("MAPI");
my $folder = $namespace->GetDefaultFolder(olFolderCalendar);
my $items = $folder->Items;

for my $itemIndex (1..$items->Count)
{
  my $message = $items->item($itemIndex);
  next if not defined $message;

  print "subject = " . $message->subject . "\n";
  print "Start         = " . $message->start . "\n";
}
