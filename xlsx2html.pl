#!/usr/bin/perl -l ../../libs/Spreadsheet-XLSX/lib/

# use lib '../../libs/Spreadsheet-XLSX/lib/';
use FindBin;
use lib "$FindBin::RealBin/libs/Spreadsheet-XLSX/lib";
#use lib "/home/www/fresta/fresta-text-conv/libs/Spreadsheet-XLSX/lib";
use strict;
use warnings;
use utf8;
use Encode;
use CGI;
use File::Basename;
use File::Slurp;
use HTML::Entities;
use HTML::Template;
use Spreadsheet::XLSX;

our $VERSION = '1.1.0';
our $DEBUG = 0;
# Directories
my $pwd = dirname($0);
#my $master_dir = $pwd.'/../xlsx';
my $master_dir = exists($ENV{MASTER_DIR}) ? $ENV{MASTER_DIR} : "/home/www/fresta/textdata/2020";
my $output_dir = $pwd;

# print header
my $cgi = CGI->new;
my $execute = $cgi->param('execute');
if( @ARGV ){
  # call from cli
  $execute = 1; 
}

if( $0 =~ /\.cgi$/ ) {
  print $cgi->header( -charset => 'utf-8' );
  print $cgi->start_html( -lang => 'ja', -encoding => 'utf-8',
    -title => "XLSX to HTML updater, Version $VERSION" );
  print "<pre>";
  print "XLSX to HTML updater, Version $VERSION";
  print "Convert files from $master_dir to $output_dir";
  if( ! $execute ) {
    print "<div><span style='border: black 2px solid; padding: 2px 8px;'><a href='?execute=1'>Convert</a></span></div>";
  }
}


print "Convert XLSX to HTML tree.\n";

opendir my $master_dh, $master_dir or die "Can't open directory $master_dir: $!";
while ( my $file = readdir $master_dh ) {
  my $infile = "$master_dir/$file";
  next if( $infile !~ /\.xlsx/ || ! -f $infile );
 
  my $dst = $file;
  $dst =~ s/\.xlsx//;
  if( $execute ) {
    # convert
    print "Convert $file to <a href='$dst/index.html'>$dst</a>";
    system("cd $master_dir; git pull");
    convert_xlsx_to_html( $file, $master_dir, $output_dir );
  }else{
    # just print a list
    print "$file will convert into <a href='$dst'>$dst</a>\n";
  }
}
if( $0 =~ /\.cgi$/ ) {
  print "done. <a href='?'>Return</a>" if $execute;
  print "</pre>";
  print "</body>";
  print "</html>";
}
exit;

sub convert_xlsx_to_html {
  my ($infile, $master_dir, $output_dir ) = @_;
  our $DEBUG;

  # source directories for files copying.
  my @src_dirs;

  my $outdir = "$output_dir/$infile";
  $outdir =~ s/\.xlsx$//;
  # create directory unless exists
  if( -e $outdir ) {
    print "Remove exist directory: $outdir" if $DEBUG;
    system ('rm','-rf',$outdir);
  }
  mkdir($outdir) unless -d $outdir;

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
  chomp($ch);
  next unless( $ch =~ /(\d+)$/ );
  my $i = $1;
  $chapters{$i} = $ch;
  $navigation{$ch}{name} = $ch;
  
  # scan description
  $sheet->{MaxRow} ||= $sheet->{MinRow};
  my $row = $sheet->{MinRow};
  for( ; $row <= $sheet->{MaxRow} ; $row++ ) {
    last unless exists( $sheet->{Cells}[$row][0]->{Val} );
    my $key = decode_entities($sheet->{Cells}[$row][0]->{Val});
    chomp($key);
    last unless( $key );
    my $val = decode_entities($sheet->{Cells}[$row][1]->{Val});
    chomp($val);
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
  printf("Sheet: %s\n", $ch) if $DEBUG;

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
    # print "INFO: $key = $val\n";
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
        my $key = sprintf('col%u',$col);

        # sanitize contents
        if( $col == 1 ) {
          my @out;
          my @lines = split(/\s*[\r\n]\s*/,$str);
          my $was_add_number = 0;
          foreach my $l ( @lines ) {
            chomp($l);
            if( $l ) { 
              if( $ch ne 'index' && @contents && ! $was_add_number && ! $rowdata{col0} && $l !~ /^\s*<h/ ) {
                $l = sprintf('[%u] ', $row_num++ ) . $l;
                $was_add_number++;
              }
              # print "D: l=$l, row_num=$row_num, was_add_number=$was_add_number\n";
              $l = '<p>'.$l.'</p>' unless $l =~ m!^\s*<!; # wrap <p> tag unless the line is wrapped by any tag manually
              push( @out, $l);
            }else{
              push( @out, '<br>') ;
            }
          }
          $str = join("\n",@out);
        }

        chomp($str);
        $rowdata{$key} = $str;
      }
    }
    push( @contents, \%rowdata);
  }

  # Outout HTML using template
  my $tmplname = exists($info{template}) ? $info{template} : '';
  unless( $tmplname ) {
    $tmplname = $infile;
    if( $tmplname =~ m!([^/]+)\.xlsx?$! ) {
      $tmplname = $1;
    }
  }
  print "Open $tmplname.tmpl as HTML Template." if $DEBUG;
  my $template = HTML::Template->new(filename=>"$master_dir/$tmplname.tmpl", die_on_bad_params=>0 );
  die("template cannot open : $tmplname.tmpl") unless $template;

  # Set parameters
  foreach my $key ( keys %info ) {
    $template->param( $key => $info{$key} );
  }
  foreach my $key ( keys %{$navigation{$ch}} ) {
    $template->param( $key => $navigation{$ch}{$key} ) unless exists( $info{$key} );
  }
  # The First content
  my $content0 = shift(@contents);
  $template->param( content0 => $$content0{col1} );

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
  push(@src_dirs,$info{template}) if( $info{template} && ! grep {$_ eq $info{template} } @src_dirs );

  # output html
  my $outfile = "$outdir/$ch.html";
  print "Output into $outfile" if $DEBUG;
  open(my $output_dh,'>',$outfile);
  $template->output(print_to => $output_dh);
  close($output_dh);

  }

  # Copy files
  my $srcdir = $infile;
  $srcdir =~ s/\.xlsx$//;
  push(@src_dirs, $srcdir) if -d $srcdir;

  foreach my $dir ( @src_dirs ) {
    my $src = "$master_dir/$dir";
    my $dst = $outdir;
    next unless -d $src;
    print "Copy files from $src to $dst\n" if $DEBUG;
    system "cp -au $src/* $dst/";
  }
}
