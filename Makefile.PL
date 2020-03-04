use 5.6.1;
use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    NAME              => 'fresta-text-conv',
    VERSION_FROM      => 'xlsx2html.pl', #'lib/FrestaTextConv/XLSX.pm', # finds $VERSION
    PREREQ_PM         => {
	    'Archive::Zip' => '1.18',
	    'Spreadsheet::ParseExcel' => '0.59',
      'HTML::TreeBuilder::XPath' => '0.14',
      'Spreadsheet::XLSX' => '0.13.1',
      'Excel::Writer::XLSX' => '1.03',
      'HTML::Template' => '2.95',
    }, # e.g., Module::Name => 1.1
    ($] >= 5.005 ?     ## Add these new keywords supported since 5.005
      ( # ABSTRACT_FROM  => 'lib/Spreadsheet/XLSX.pm', # retrieve abstract from module
       AUTHOR         => 'Yuji Tsuchimoto <yuji@eikai.co.jp>' ) : ()),
);
