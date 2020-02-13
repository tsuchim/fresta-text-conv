#!/usr/bin/perl

use lib 'libs/Spreadsheet-XLSX/lib/';
use strict;
use warnings;
use utf8;
use Encode;
use File::Slurp;
use HTML::Entities;
use HTML::Template;
use Spreadsheet::XLSX;


# Open book file
my $infile = 'data/xlsx/win10.xlsx';
my $excel = Spreadsheet::XLSX->new($infile);
 
# show book info
#print "Filename :", $excel->{File} , "\n";
#print "Sheet Count :", $excel->{SheetCount} , "\n";
#print "Author:", $excel->{Author} , "\n";

# show sheet and cell info
foreach my $sheet (@{$excel -> {Worksheet}}) {
  printf("Sheet: %s\n", $sheet->{Name});

  # Read Information from a Header of the Sheet
  my %info;
  $sheet->{MaxRow} ||= $sheet->{MinRow};
  my $row = $sheet->{MinRow};
  for( ; $row <= $sheet->{MaxRow} ; $row++ ) {
    my $key = decode_entities($sheet->{Cells}[$row][0]->{Val});
    last unless( $key );
    my $val = decode_entities($sheet->{Cells}[$row][1]->{Val});
    $info{$key} = decode_entities($val);
    print "INFO: $key = $val\n";
  }

  # Read Contents from the Sheet
  $row++;
  for( ; $row <= $sheet->{MaxRow} ; $row++ ) {
    $sheet -> {MaxCol} ||= $sheet -> {MinCol};
    foreach my $col ($sheet -> {MinCol} ..  $sheet -> {MaxCol}) {
      my $cell = $sheet -> {Cells} [$row] [$col];
      if ($cell) {
        my $str = decode_entities( $cell->{Val});
        printf("( %s , %s ) => %s\n", $row, $col, $str);
      }
    }
  }

  # Outout HTML using Templete 
}
