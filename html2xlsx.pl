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
  if( $env_dir !~ /^\./ && -d $fulldir ) {
    print "$fulldir as $env_dir\n";
    push( @text_envs, $env_dir);
  }
}

# chapter ごとにファイルを作成
for( my $ch = 0 ; $ch <= 6 ; $ch++ ) {
  # HTML から必要な情報を読み取ってエクセルにまとめる
  my $testfile = $master_dir.$text_envs[0]."/ch$ch.html";
  next unless -f $testfile;

  # Create XLS file
  my $xlsxfile = "data/xlsx/ch$ch.xlsx";
  print "Open $xlsxfile\n";
  my $workbook = Excel::Writer::XLSX->new($xlsxfile);
  # Set default font
  #my $font_name = ''; # decode("cp932","游ゴシック");
  #workbook->{_formats}->[15]->set_properties(font  => $font_name, size  => 11, align => 'vcenter');

  # Scan all environments
  foreach my $env_dir ( @text_envs ) {
    my $infile = $master_dir.$env_dir."/ch$ch.html";
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
    my $worksheet = $workbook->add_worksheet($env_dir);

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