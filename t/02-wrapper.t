use strict;
use warnings;

use Test::Fatal;
use Test::More 0.98;

use Excel::CloneXLSX::WrappedParser;
use Spreadsheet::ParseExcel::Utility qw(sheetRef);


my $wrapped = 't/data/wrapped.xlsx';

{
    like exception {
        Excel::CloneXLSX::WrappedParser->new({filehandle => $wrapped})
      }, qr{did not pass type constraint.+filehandle}i,
          "filehandle must be a filehandle";

}


{
    open my $fh, '<', $wrapped or die "Can't open $wrapped: $!";
    my $parser = Excel::CloneXLSX::WrappedParser->new({filehandle => $fh});
    $parser->workbook;

    my ($worksheet, $red, $green, $blue)
        = ('EmptyCells', '#ff0000', '#008000', '#0000ff');
    my @bgcolors = (
        ['A1', '' ],
        ['B1', $blue ],
        ['C1', $green],
        ['A2', $red  ],
        ['B2', $blue ],
        ['C2', $green],
        ['A3', $green],
        ['B3', $red  ],
        ['C3', $green],
        ['D4', $green],
    );

    for my $bgcolor (@bgcolors) {
        my ($cell, $color) = @$bgcolor;
        my $fmt = $parser->get_formatting_for_cell( $worksheet, sheetRef($cell) );
        is lc($fmt->{Fill}[1] || ''), $color, "Cell $cell has bgcolor $color";
    }
}


done_testing;
