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
my $tmpldir = 'data/xlsx';
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
  my @contents;
  for( ; $row <= $sheet->{MaxRow} ; $row++ ) {
    $sheet->{MaxCol} ||= $sheet->{MinCol};
    my %rowdata;
    foreach my $col ($sheet->{MinCol} ..  $sheet->{MaxCol}) {
      my $cell = $sheet->{Cells}[$row][$col];
      if ($cell) {
        my $str = decode_entities( $cell->{Val});
        # printf("( %s , %s ) => %s\n", $row, $col, $str);
        my $key = 'class';
        $key = sprintf('col%u',$col) if( $col );
        $rowdata{$key} = $str;
      }
    }
    push( @contents, \%rowdata);
  }

  # Outout HTML using Templete
  my $tmplname = exists($info{templete}) ? $info{templete} : '';
  unless( $tmplname ) {
    $tmplname = $infile;
    if( $tmplname =~ m!([^/]+)\.xlsx?$! ) {
      $tmplname = $1;
    }
  }
  print "Open $tmplname.tmpl as HTML Template.\n";
  my $template = HTML::Template->new(filename=>"$tmpldir/$tmplname.tmpl", die_on_bad_params=>0 );
  die("Templete cannot open : $tmplname.tmpl") unless $template;

  # set parameters
  foreach my $key ( keys %info ) {
    $template->param( $key => $info{$key} );
  }
  $template->param( contents0 => shift(@contents) );
  if( $sheet->{Name} eq 'index' ) {
    # Set parameters of contents individually for index
    my $i=0;
    foreach my $content ( @contents ) {
      $i++;
      $template->param( "content$i" => $content );
    }
  }else{
    # Set parameters of contents as array for normal contents
    $template->param( contents => \@contents );
  }
  # Output
  print $template->output();
}
