use 5.008004;
use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    NAME              => 'Win32::Ecco::Import::OML',
    VERSION_FROM      => 'lib/Win32/Ecco/Import/OML.pm', # finds $VERSION
    PREREQ_PM         => {}, # e.g., Module::Name => 1.1
    ($] >= 5.005 ?     ## Add these new keywords supported since 5.005
      (ABSTRACT       => 'Import from OML to Ecco', # retrieve abstract from module
       AUTHOR         => 'Will Sargent <will_sargent@yahoo.com>') : ()),
);
