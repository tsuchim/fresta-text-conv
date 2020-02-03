#!/usr/bin/perl

use lib 'libs/Spreadsheet-XLSX/lib/';
use strict;
use warnings;
use utf8;
use Encode;
use File::Slurp;
use HTML::TreeBuilder;
use Spreadsheet::XLSX;


# Open book file
my $infile = 'data/chapter-0.xlsx';
my $oBook = Spreadsheet::XLSX->new($infile);
 
# show book info
print "Filename :", $oBook->{File} , "\n";
print "Sheet Count :", $oBook->{SheetCount} , "\n";
print "Author:", $oBook->{Author} , "\n";
 
# show sheet and cell info
my $oSheet = $oBook->{Worksheet}[0];
print "Sheet Name:" , $oSheet->{Name} , "\n";
print "A1:", $oSheet->{Cells}[0][0]->value, "\n";