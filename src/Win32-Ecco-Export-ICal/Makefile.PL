use 5.008004;
use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    NAME              => 'Win32::Ecco::Export::ICal',
    VERSION_FROM      => 'lib/Win32/Ecco/Export/ICal.pm', # finds $VERSION
    PREREQ_PM         => { Data::ICal => 0.05 }, # e.g., Module::Name => 1.1
    ($] >= 5.005 ?     ## Add these new keywords supported since 5.005
      (ABSTRACT       => 'Exports from Ecco to ICal',
       AUTHOR         => 'Will Sargent <will_sargent@yahoo.com>') : ()),
);
