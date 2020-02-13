#!/usr/bin/perl

use lib '../../libs/Spreadsheet-XLSX/lib/';
use strict;
use warnings;
use utf8;
use Encode;
use CGI;
use File::Slurp;
use HTML::Entities;
use HTML::Template;
use Spreadsheet::XLSX;

our $VERSION = '0.10';

# Directories
my $master_dir = '../xlsx';
my $output_dir = '.';

# print header
my $cgi = CGI->new;
my $execute = $cgi->param('execute');

if( $0 =~ /\.cgi$/ ) {
  print $cgi->header();
  print "<html>";
  print "<head><title>XLSX to HTML updater, Version $VERSION</title></head>";
  print "<body>";
  print "<pre>";
}

if( @ARGV ){
  # call from cli
  $execute = 1; 
}

print "Convert XLSX to HTML tree.\n";

opendir my $master_dh, $master_dir or die "Can't open directory $master_dir: $!";
while ( my $file = readdir $master_dh ) {
  my $infile = "$master_dir/$file";
  next if( $infile !~ /\.xlsx/ || ! -f $infile );
 
  if( $execute ) {
    # convert
    convert_xlsx_to_html( $file, $master_dir, $output_dir );
  }else{
    # just print a list
    print "$file\n";
  }
}
if( $0 =~ /\.cgi$/ ) {
  print "</pre>";
  print "</body>";
  print "</html>";
}
exit;

sub convert_xlsx_to_html {
  my ($infile, $master_dir, $output_dir ) = @_;

  my $excel = Spreadsheet::XLSX->new("$master_dir/$infile");
 
# show book info
#print "Filename :", $excel->{File} , "\n";
#print "Sheet Count :", $excel->{SheetCount} , "\n";
#print "Author:", $excel->{Author} , "\n";

# scan sheets and create navigation
my %chapters; # list of chapter number => chapter name
my $chapter_start = 1; # the first chapter number
my %navigation; # lists of infomation for navigation and index for
my @index; # lists of chapters for index
foreach my $sheet (@{$excel->{Worksheet}}) {
  my $ch = $sheet->{Name};
  next unless( $ch =~ /\d+$/ );
  my $i = $&;
  $chapters{$i} = $ch;
  $navigation{$ch}{name} = $ch;
  
  # scan description
  $sheet->{MaxRow} ||= $sheet->{MinRow};
  my $row = $sheet->{MinRow};
  for( ; $row <= $sheet->{MaxRow} ; $row++ ) {
    last unless exists( $sheet->{Cells}[$row][0]->{Val} );
    my $key = decode_entities($sheet->{Cells}[$row][0]->{Val});
    last unless( $key );
    my $val = decode_entities($sheet->{Cells}[$row][1]->{Val});
    $navigation{$ch}{$key} = $val;
  }
}
foreach my $i ( sort keys %chapters ) {
  my $ch = $chapters{$i};
  $chapter_start = $i if( $i < $chapter_start ); # the start number of chapter (ch0 exists on some envs)
  $navigation{$ch}{prev} = $chapters{$i-1} if exists($chapters{$i-1});
  $navigation{$ch}{next} = $chapters{$i+1} if exists($chapters{$i+1});
  $navigation{$ch}{title} = $navigation{$ch}{header1};
  push( @index, $navigation{$ch} );
}

# show sheet and cell info
my $ch_num = $chapter_start - 1;
foreach my $sheet (@{$excel->{Worksheet}}) {
  my $ch = $sheet->{Name};
  $ch_num++;
  printf("Sheet: %s\n", $ch);

  # Read Information from a Header of the Sheet
  my %info;
  $info{chapter_start} = $chapter_start;
  $info{chapter_first} = $chapters{$chapter_start};

  $sheet->{MaxRow} ||= $sheet->{MinRow};
  my $row = $sheet->{MinRow};
  for( ; $row <= $sheet->{MaxRow} ; $row++ ) {
    last unless exists($sheet->{Cells}[$row][0]->{Val});
    my $key = decode_entities($sheet->{Cells}[$row][0]->{Val});
    last unless( $key );
    my $val = decode_entities($sheet->{Cells}[$row][1]->{Val});
    $info{$key} = decode_entities($val);
    print "INFO: $key = $val\n";
  }
  $info{header1} = sprintf('%u. %s', $ch_num-1, $info{header1}) if exists($info{header1}) && 0 < $ch_num; # add chapter number to header1

  # Read Contents from the Sheet
  $row++;
  my $row_num = 1;
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
        # sanitize contents
        if( $col == 1 ) {
          my @out;
          my @lines = split(/\s*[\r\n]\s*/,$str);
          foreach my $l ( @lines ) {
            chomp($l);
            if( $l ) {
              $l = sprintf('[%u] ', $row_num++ ) . $l if( $ch ne 'index' && @contents && ! $rowdata{class} && ! @out );
              $l = '<p>'.$l.'</p>' unless $l =~ m!^<!; # wrap <p> tag unless the line is wrapped by any tag manually
              push( @out, $l);
            }else{
              push( @out, '<br>') ;
            }
          }
          $str = join("\n",@out);
        }
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
  my $template = HTML::Template->new(filename=>"$master_dir/$tmplname.tmpl", die_on_bad_params=>0 );
  die("Templete cannot open : $tmplname.tmpl") unless $template;

  # Set parameters
  foreach my $key ( keys %info ) {
    $template->param( $key => $info{$key} );
  }
  foreach my $key ( keys %{$navigation{$ch}} ) {
    $template->param( $key => $navigation{$ch}{$key} ) unless exists( $info{$key} );
  }
  $template->param( contents0 => shift(@contents) );

  # Set contents
  if( $sheet->{Name} eq 'index' ) {
    # Set parameters of contents individually for index
    my $i=0;
    foreach my $content ( @contents ) {
      $i++;
      $template->param( "content$i" => $$content{col1} );
    }
    $template->param( index => \@index );
  }else{
    # Set parameters of contents as array for normal contents
    $template->param( contents => \@contents );
  }

  # Output
  my $outfile = "$output_dir/$infile";
  $outfile =~ s/\.xlsx$//;
  # create directory unless exists
  mkdir($outfile) unless -d $outfile;
  # output html
  $outfile .= "/$ch.html";
  print "Output into $outfile\n";
  open(my $output_dh,'>',$outfile);
  $template->output(print_to => $output_dh);
  close($output_dh);
}
}