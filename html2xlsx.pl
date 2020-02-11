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
  # Define Formats
  my $format_wrap = $workbook->add_format();
  $format_wrap->set_text_wrap();

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

    $info{title} = $_ for $tree->findnodes(q{//head/title});
    $info{header1} = $_ for $tree->findnodes(q{//body//h1});
    $info{header1} =~ s!>\d\.\s*!>!; # チャプター番号を削る

    push( @contents, $tree->findnodes(q{//body//div[@id="header"]/div[@class="container"]}) );
    push( @contents, $tree->findnodes(q{//body//div[@class="row"]}) );

    # Add a worksheet
    my $worksheet = $workbook->add_worksheet($ch);
    #SetWidth
    $worksheet->set_column(1,1,100,$format_wrap);
    $worksheet->set_column(2,2,20);

    my $row = 0;
    foreach my $key ( keys %info ) {
      my @strs = extract_header_from_node($info{$key});
      $worksheet->write_string( $row, 0, $key );
      for( my $col=0 ; $col<@strs ; $col++ ) {
        $worksheet->write_string( $row, $col+1, decode('utf8', $strs  [$col] ) );
      }
      print "$key => ".join(',',@strs)."\n";
      $row++;
    }
    $row++;
    foreach my $content ( @contents ) {
      my @strs = extract_contents_from_node($content);
      next unless $strs[1]; # コンテンツがない場合は飛ばす
      for( my $col=0 ; $col<@strs ; $col++ ) {
        $worksheet->write_string( $row, $col, decode('utf8',$strs[$col]) );
      }
      print "$row : ".join(',',@strs)."\n";
      $row++;
    }
  }
  # Close
  $workbook->close();
}

sub extract_header_from_node {
  my $node = shift(@_);
  $_ = $node->as_text;
  return ($_);
}
sub extract_contents_from_node {
  my $node = shift(@_);
  my $class = '';
  $_ = $node->as_XML;
  my $image = '';
  # ヘッダーのパターン
  if( m!<div\s+class="container">(.*?)</div>! ) {
    $_ = $1;
    s!<h1.*?</h1>!!;
  }
  # ナビパターン
  if( m!<div[^<>]*id="navsp"! ) {
    $_ = '';
  }
  # 本文と画像のパターン
  if( m!<div class="row">(.+)</div>! ) {
    $_ = $1;
    # 画像を抽出
    if( s!<div[^<>]+>\s*<img[^<>]+?src="image/([^"]+)"[^<>]*>\s*</div>!!s ) {
      $image = $1;
    }

    # 注意パネルケース
    if( s!<div[^<>]+?class="(panel[^"]+)[^<>]+>(.+)</div>!$2!s ) {
      $class = $1;
    }

    # 本文を抽出
    s!<div[^<>]*>\s*(<p>)\s*\[\d+\]\s*(.+)</div>!$1$2!s unless( $class );
  }

  # <p>を展開
  s!<p>(.*?)</p>!$1\n!gism;
  # 改行を展開
  s!<[^<>]*br[^<>]*>!<br>\n!gim;
  # 終了タグで改行
  s!</(?:ol|ul||li|h\d)[^<>]*>!$&\n!gim;
  # 値を返す
  return ($class,$_,$image);
}

