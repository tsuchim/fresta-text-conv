#!/usr/bin/perl

use lib 'libs/Excel-Writer-XLSX/lib';
use strict;
use warnings;
use utf8;
use Encode;
use File::Slurp;
use XML::XPathEngine;
use HTML::TreeBuilder::XPath;
use Excel::Writer::XLSX;

# ディレクトリツリーを辿る
my $master_dir = 'data/fresta-text-2019/2019/';
opendir my $master_dh, $master_dir or die "Can't open directory $master_dir: $!";

# 環境一覧(ディレクトリ一覧)を取得
my @text_envs;
while ( my $env_dir = readdir $master_dh ) {
  my $fulldir = "$master_dir$env_dir";
  next if( $env_dir =~ /^\./ || ! -d $fulldir );
  print "$fulldir as $env_dir\n";

  # ディレクトリごとにエクセルファイルを作る
  my $xlsxfile = "data/xlsx/$env_dir.xlsx";
  print "Open $xlsxfile\n";
  my $workbook = Excel::Writer::XLSX->new($xlsxfile);
  # Set default font
  #my $font_name = ''; # decode("cp932","游ゴシック");
  #workbook->{_formats}->[15]->set_properties(font  => $font_name, size  => 11, align => 'vcenter');

  # chapter ごとにシートを作成
  my @chapters = ('index');
  for( my $i = 0 ; $i <= 9 ; $i++ ) {
    # HTML から必要な情報を読み取ってエクセルにまとめる
    my $testfile = $master_dir.$env_dir."/ch$i.html";
    next unless -f $testfile;
    push( @chapters, "ch$i");
  }

  # Scan all environments
  foreach my $ch ( @chapters ) {
    my $infile = $master_dir.$env_dir."/$ch.html";
    unless( -f $infile ) {
      print "Skip $infile, file not found.\n";
      next;
    }
    my $html = read_file($infile);
    my $tree = HTML::TreeBuilder::XPath->new;
    $tree->parse($html);

    # 必要な情報を配列にストア
    my %info;
    my @contents;

    $info{title} = $_ for $tree->findnodes_as_strings(q{//head/title});
    $info{header1} = $_ for $tree->findnodes_as_strings(q{//body//h1});

    push( @contents, $tree->findnodes_as_strings(q{//body//div[@id="header"]}) );
    push( @contents, $tree->findnodes_as_strings(q{//body//div[@class="row"]}) );

    # Add a worksheet
    my $worksheet = $workbook->add_worksheet($ch);

    my $row = 0;
    foreach my $key ( keys %info ) {
      $worksheet->write_string( $row, 0, $key );
      $worksheet->write_string( $row, 1, decode('utf8',$info{$key}) );
      print "$key => $info{$key}\n";
      $row++;
    }
    $row++;
    foreach my $content ( @contents ) {
      $worksheet->write_string( $row, 1, decode('utf8',$content) );
      $row++;
    }
  }
  # Close
  $workbook->close();
}