This library converts OML to and from Ecco format, using Windows DDE.

The perl scripts may not make all that much sense to you.  It helps
if you've read Dave Allen's "Getting Things Done" as it explains the 
configuration of the file names and how Ecco is organized.  

The way that the perl scripts are broken down:

ecco.pm -- this is the main Ecco library.  It contains all the 
DDE specific functionality, and makes it look like there's a single 
cohesive Ecco object you can manipulate using the Ecco API (specified in 
api.eco).

export.pl -- this is a wrapper script which takes some arguments from outside, and 
calls the appropriate export script (that is, exportActions, exportProjects, or 
exportSomeday) to create an OML file, and then takes that OML file and generates a 
Shadow XML file from it.

  export.pl "Actions"

import.pl -- this script parses out the OML file given and imports it into 
the Ecco folder noted on the command line.


  import.pl "folderName" "fileName"  

inbasket.pl -- this script calls the import.pl script for the Inbasket items in shadow,
and then sets the Shadow items to be deleted.  This means that those items are removed
from the Handheld device and appear in Ecco.

command.pl -- this script allows the text on the command line to be imported as a 
command.

diary.pl -- does the same thing as inbasket.pl, but is for personal diary entries.

sync.pl -- If items in Shadow are marked as checked, then marks the matching item in Ecco as
done, with Shadow's Finish Date.

syncall.pl -- Calls all of the appropriate scripts.

google.pl -- Looks up Google with the selected item text.



The way I deal with this is basically:

1. New stuff I have to do is added to the Palm Inbasket.
2. When I sync, the ShadP-Inbasket.XML file is updated.
3. When I import, those items are sucked into the Ecco Inbasket folder.
4. I set the appropriate context and action on the items.
5. I then export the Action, Projects and Someday folders from Ecco, back to the XML.
6. When I do a sync, the Inbasket on the Palm is cleared, and the new exported Actions, 
   Projects and Someday files are shown on the Palm.

You must have the following ActivePerl modules to run the scripts:

XML-DOM -- this is required for parsing the Shadow and OML XML.
TimeDate -- used for converting dates between formats
DDE -- This is a custom windows DDE found at http://www.bribes.org/perl/ppm/Win32-DDE.ppd



