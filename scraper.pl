#!/usr/bin/perl

use strict;
use warnings;

use URI;
use Web::Scraper;
use String::ShellQuote;
use Excel::Writer::XLSX;
use Set::Light;
use Encode;

my $doc_scraper = scraper {
	process 'h4 > a', 'links[]' => '@href';
	process 'h4', 'names[]' => 'TEXT';
};

my $result = $doc_scraper->scrape(URI->new('http://www.doe.mass.edu/cte/frameworks/?section=all'));

foreach my $link (@{$result->{links}})  {
	# Make sure some 1337 hax0r doesn't think it's funny to change the link to $(rm -rf ~/*).pdf
	$link = shell_quote($link);
	# Perl is magic holy crap you can do this in one line
	if ($link =~ /([^\.\/]*)\.pdf$/) {
		open(my $fh, "wget $link -O - | pdftotext -layout - - |");
		binmode($fh, 'encoding(UTF-8)');
		file_to_spreadsheet($1, $fh, shift(@{$result->{names}})); #this is a "bad" idea, but it "works" :^)
		undef($fh);
	}
}

sub set_column_widths
{
	my ($workbook) = @_;
	my $sheet = $workbook->sheets(0);
	$sheet->set_column('B:E', 2);
	$sheet->set_column('G:G', 70);
	$sheet->set_column('F:F', 10);
	$sheet->set_column('I:I', 60);
}

sub make_header
{
	my ($workbook, $title) = @_;
	my $sheet = $workbook->sheets(0);
	my $title_format = $workbook->add_format(
		font => 'Arial',
		color => 'blue',
		size => 24,
		bold => 1,
		align => 'center',
		valign => 'top',
		text_wrap => 1
	);
	$sheet->insert_image('B2', './qhshat.png', 0, 0, 0.4, 0.4);
	$sheet->write('G2', $title, $title_format);
	$sheet->set_row(1, 100);

	my $bold_format = $workbook->add_format(
		bold => 1,
		font => 'Arial',
		size => 10,
		valign => 'top'
	);
	$sheet->write('B3', 'Q1', $bold_format);
	$sheet->write('C3', 'Q2', $bold_format);
	$sheet->write('D3', 'Q3', $bold_format);
	$sheet->write('E3', 'Q4', $bold_format);

	my $format = $workbook->add_format(
		font => 'Arial',
		size => 10,
		valign => 'top'
	);

	$sheet->write('I2', "This sheet was automatically generated by a script:\nhttps://github.com/blucoat/MassFrameworks\nIf you find any errors, please contact James Hagborg<jameshagborg\@gmail.com>.", $format);
}

sub file_to_spreadsheet
{
	my ($name, $fh, $title) = @_;
	my $workbook = Excel::Writer::XLSX->new("$name.xlsx");
	$workbook->add_worksheet();
	my $sheet = $workbook->sheets(0);

	make_header($workbook, $title);
	set_column_widths($workbook);

	my $format = $workbook->add_format(
		font => 'Arial',
		size => 10,
		valign => 'top'
	);
	my $bold_format = $workbook->add_format(
		bold => 1,
		font => 'Arial',
		size => 10,
		valign => 'top'
	);

	# Actually parse the data out
	# This feels hacky, but it's the most reliable thing I could come up with
	my $indent = 0;
	my $started = 0;
	my $current_str = "";
	my $height = 0;
	my $row = 4;
	my $previd = "";
	my $idset = Set::Light->new();
	while (my $line = <$fh>) {
		if ($line =~ /^[\s]*Strand [0-9]+:/) {
			next;
		} elsif ($line =~ /^[\s]*Appendices[\s]*$/ and $started) {
			last;
		} elsif ($line =~ /^[\s]*(([0-9])\.[A-Z][^\s]*)[\s]*(.*)/ and ($started or $2 eq "1") and not $idset->contains($1)) {
			my $id = $1;
			my $desc = $3;
			$indent = $-[3];

			$idset->insert($1 =~ s/\*//gr);

			if ($previd and $previd =~ /[A-Z]$/) {
				$sheet->write("G$row", $current_str, $bold_format);
			} else {
				$sheet->write("G$row", $current_str, $format);
			}


			$sheet->set_row($row - 1, $height * 12);
			$row++;
			$sheet->write("F$row", "$id", $format);

			$height = 1;
			$started = 1;
			$current_str = $desc;
			$previd = $1;
		} elsif ($indent != 0 and $line =~ /^[\s]*(.*)/ and abs($-[1] - $indent) < 4) {
			$height++;
			$current_str .= "\n" . $1;
		}
	}
	#All done, finish last row
	$sheet->write("G$row", $current_str, $format);
	$sheet->set_row($row - 1, $height * 12);
	$workbook->close();
}

