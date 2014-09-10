#!/usr/bin/perl

# Copyright 2014 James Hagborg
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

use strict;
use warnings;

use URI;
use Web::Scraper;
use String::ShellQuote;
use Excel::Writer::XLSX;
use Set::Light;
use Encode;

# Declare global fonts, these get updated when a new worksheet is made
# A better solution would be OO, but I'm not familiar with that aspect of perl yet
my $title_format;
my $bold_format;
my $bold_format_noline;
my $text_format;
my $name_format;
my $subtitle_format;
my $qhs_format;

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
		my $name = $1;
		my $title = shift(@{$result->{names}});
		if (not $title =~ /[\s]*([^\(]*)/) {
			close($fh);
			next;
		}
		$title = $1;
		file_to_spreadsheet($name, $fh, $title);
		undef($fh);
	}
}

sub set_column_widths
{
	my ($workbook) = @_;
	my $sheet = $workbook->sheets(0);
	$sheet->set_column('A:D', 1);
	$sheet->set_column('F:F', 60);
	$sheet->set_column('E:E', 10);
}

sub make_fonts
{
	my ($workbook) = @_;

	$title_format = $workbook->add_format(
		font => 'Cambria',
		size => 20,
		bold => 1,
		align => 'left',
		valign => 'top',
		text_wrap => 1
	);

	$subtitle_format = $workbook->add_format(
		font => 'Cambria',
		size => 12,
		bold => 1,
		align => 'left',
		valign => 'top',
		text_wrap => 1
	);

	$qhs_format = $workbook->add_format(
		font => 'Cambria',
		color => 'blue',
		size => 14,
		bold => 1,
		align => 'left',
		valign => 'top',
		text_wrap => 1
	);

	$name_format = $workbook->add_format(
		font => 'Cambria',
		size => 12,
		bold => 1,
		align => 'left',
		valign => 'top',
		text_wrap => 1
	);

	$bold_format = $workbook->add_format(
		bold => 1,
		font => 'Times New Roman',
		size => 10,
		valign => 'top',
		align => 'left',
		border => 1
	);

	$bold_format_noline = $workbook->add_format(
		bold => 1,
		font => 'Times New Roman',
		size => 10,
		valign => 'top',
		align => 'left',
	);

	$text_format = $workbook->add_format(
		font => 'Times New Roman',
		size => 10,
		valign => 'top',
		align => 'left',
		border => 1
	);
}

sub make_header
{
	my ($workbook, $title, $subtitle) = @_;
	my $sheet = $workbook->sheets(0);

	#set to letter paper
	#for some reason this commie software defaults to A4
	$sheet->set_paper(1);
	#live life on the edge :^)
	$sheet->set_margins_LR(0.4);

	$sheet->insert_image('A1', './qhshat.png', 0, 0, 0.3, 0.3);
	$sheet->write('F1', $title, $title_format);
	$sheet->write('F2', $subtitle, $subtitle_format);
	$sheet->write('F3', 'Quincy High School	       School Year: __________', $qhs_format);
	$sheet->write('F4', 'Student Name: ______________________________________', $name_format);

	$sheet->set_row(0, 50);
	$sheet->set_row(1, 30);
	$sheet->set_row(2, 25);
	$sheet->set_row(3, 20);

	$sheet->write('A5', '1', $bold_format);
	$sheet->write('B5', '2', $bold_format);
	$sheet->write('C5', '3', $bold_format);
	$sheet->write('D5', '4', $bold_format);

	#do this to fill in the borders
	$sheet->write("E5", '', $text_format);
	$sheet->write("F5", '', $text_format);
}

sub make_footer
{
	my ($sheet, $row) = @_;
	$row++;
	$sheet->merge_range("A$row:F$row", 'Certifications:', $bold_format_noline);
	$row += 2;
	$sheet->merge_range("A$row:F$row", 'Internships:', $bold_format_noline);
}

sub write_strand
{
	my ($sheet, $row, $text) = @_;
	$sheet->merge_range("E$row:F$row", "$text", $bold_format);

	#do this to fill in the borders
	$sheet->write("A$row", '', $text_format);
	$sheet->write("B$row", '', $text_format);
	$sheet->write("C$row", '', $text_format);
	$sheet->write("D$row", '', $text_format);
}

sub write_row
{
	my ($sheet, $row, $id, $str, $height) = @_;

	if ($id =~ /[A-Z]\*?$/) {
		$sheet->write("F$row", $str, $bold_format);
	} else {
		$sheet->write("F$row", $str, $text_format);
	}

	#rows here count from 0, but in all other cases from 1
	$sheet->set_row($row - 1, $height * 14);
	$sheet->write("E$row", "$id", $text_format);

	#do this to fill in the borders
	$sheet->write("A$row", '', $text_format);
	$sheet->write("B$row", '', $text_format);
	$sheet->write("C$row", '', $text_format);
	$sheet->write("D$row", '', $text_format);
}

sub file_to_spreadsheet
{
	my ($name, $fh, $title) = @_;
	my $workbook = Excel::Writer::XLSX->new("$name.xlsx");
	$workbook->add_worksheet();
	my $sheet = $workbook->sheets(0);

	$sheet->set_header("&C$title");
	$sheet->set_footer("&C1-EXPOSURE 2-COMPETENT 3-PROFICIENT 4-ADVANCED\n&P");

	make_fonts($workbook);
	set_column_widths($workbook);

	# Actually parse the data out
	# This feels hacky, but it's the most reliable thing I could come up with
	my $indent = 0;
	my $started = 0;
	my $current_id = "";
	my $current_str = "";
	my $height = 0;
	my $row = 6;
	my $idset = Set::Light->new();
	my $header_made = 0;
	while (my $line = <$fh>) {
		if ($line =~ /^[\s]*(Strand ([0-9]+):([\s]?[^0-9\s])*)$/ and $2 ne '3') {
			if ($current_str) {
				write_row($sheet, $row, $current_id, $current_str, $height);
				$current_str = '';
				$indent = 0;
				$row++;
			}
			write_strand($sheet, $row, $1);
			$row++;
		} elsif ($line =~ /Performance Examples?:/) {
			if ($current_str) {
				write_row($sheet, $row, $current_id, $current_str, $height);
				$current_str = '';
				$indent = 0;
				$row++;
			}
			#skip this line, resetting stuff too
		} elsif ($line =~ /^[\s]*Appendices[\s]*$/ and $started) {
			#we've hit the end
			last;
		} elsif ($line =~ /^[\s]*(([0-9])\.[A-Z][^\s]*)[\s]*(.*)/ and ($started or $2 eq "1") and not $idset->contains($1)) {
			if ($current_str) {
				write_row($sheet, $row, $current_id, $current_str, $height);
				$row++;
			}

			$current_id = $1;
			$current_str = $3;
			$indent = $-[3];
			$height = 1;
			$started = 1;
			my $nostar = $1 =~ s/\*//gr;
			$idset->insert($nostar);
			$idset->insert($nostar . '*');
		} elsif ($indent != 0 and $line =~ /^[\s]*(.*)/ and abs($-[1] - $indent) < 4) {
			$height++;
			$current_str .= "\n" . $1;
		} elsif ($line =~ /[\s]*(.*Cluster)[\s]*$/) {
			make_header($workbook, $title, $1);	
			$header_made = 1;
		}
	}
	#All done, finish last row
	write_row($sheet, $row, $current_id, $current_str, $height);
	make_header($workbook, $title, '') if (not $header_made);
	make_footer($sheet, $row);
	$workbook->close();
}

