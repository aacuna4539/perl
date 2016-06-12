#!/usr/bin/perl -w
use strict;
use warnings;
use Data::Printer;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;

my $path = "";
my $filename = "";
my $worksheet_idx = "";
    
my $parser   = Spreadsheet::ParseExcel::SaveParser->new();
my $template = $parser->Parse($path . $filename);

# ---- TODO: Grab range, iterate, and update vals in cells
# Change to select sheet based on name (typically "Month Day Invoice")
my $worksheet = $template->worksheet($worksheet_idx); 
my $row  = 15;
my $col  = 0;

# ---- TODO: Don't hardcode val, get from user
$worksheet->AddCell( $row, $col, '15' );

# Write over the existing file or write a new file.
$template->SaveAs($path . "new " . $filename );

# ---- TODO: Formulas must be added to new file -- SaveParser cannot pull them from original file.

#my $spreadsheet = $template->worksheet(20);
#my $spreadsheet_name = $spreadsheet->get_name();
#my $workbook = $template->SaveAs($path . "new " . $filename );
#my $ws = $workbook->sheets(21);
#$workbook->close();
