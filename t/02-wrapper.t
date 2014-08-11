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
        ['A1', '',     '',    ],
        ['B1', '',     $blue, ],
        ['C1', $green, $green ],
        ['A2', '',     $red,  ],
        ['B2', $blue,  $blue, ],
        ['C2', $green, $green,],
        ['A3', $green, $green,],
        ['B3', '',     $red,  ],
        ['C3', $green, $green,],
        ['D4', $green, $green,],
    );

    for my $bgcolor (@bgcolors) {
        my ($cell, $act_color, $comp_color) = @$bgcolor;

        my $act_fmt = $parser->get_cell_format( $worksheet, sheetRef($cell) );
        is lc($act_fmt->{Fill}[1] || ''),
            $act_color, "Cell $cell has actual bgcolor $act_color";

        my $comp_fmt = $parser->get_computed_cell_format( $worksheet, sheetRef($cell) );
        is lc($comp_fmt->{Fill}[1] || ''),
            $comp_color, "Cell $cell has computed bgcolor $comp_color";
    }
}


done_testing;
