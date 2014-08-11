use strict;
use warnings;

use lib q{t/lib};

use t::CloneXLSX::Utils qw(:all);
use Test::Fatal;
use Test::More 0.98;

use Excel::CloneXLSX;
use File::Temp qw(tempfile);

my $clone_easy = 't/data/clone-easy.xlsx';
{
    my ($fh, $tempfile) = tempfile();
    Excel::CloneXLSX->new({
        from => $clone_easy,
        to   => $tempfile,
    })->clone;

    open my $new_fh, '<', $tempfile or die "Can't open $tempfile: $!";
    my $parser = Excel::CloneXLSX::WrappedParser->new({
        filehandle => $new_fh,
    });
    my $workbook = $parser->workbook;

    is_deeply [map {$_->get_name} $workbook->worksheets()],
        [qw(EmptyCells TextCells)],
            'same workbooks';

    subtest 'EmptyCells worksheet' => sub {
        my $wkst = $workbook->worksheet('EmptyCells');
        worksheet_range_is($wkst, [0,3], [0,3]);

        # green cells have cell-formatting and are q{}
        # blank, red, and blue cells have default formatting and are undef
        # EXCEPTION: B2 [1,1] the blue of the column format overrides the
        # red of the row format.  Since Row beats Col, the blue must be
        # stored as a cell-format. Ergo the cell is q{}
        cell_contents_are($wkst, [
            [undef, undef,   q{},   undef, ],
            [undef, q{},     q{},   undef, ],
            [q{},   undef,   q{},   undef, ],
            [undef, undef,   undef, q{},   ],
        ]);
        my ($red, $green, $blue) = ('#ff0000', '#008000', '#0000ff');
        cell_bgcolors_are($parser, $wkst, [
            [undef,  $blue, $green, undef,  ],
            [$red,   $blue, $green, $red,   ],
            [$green, $red,  $green, $red,   ],
            [undef,  $blue, $blue,  $green, ],

        ]);
    };

    subtest 'TextCells worksheet' => sub {
        my $wkst = $workbook->worksheet('TextCells');
        worksheet_range_is($wkst, [0,2], [0,2]);
        cell_contents_are($wkst, [ [qw(A B C)], [qw(D E F)], [qw(G H I)] ]);
        cell_bgcolors_are($parser, $wkst, [([(undef) x (3)]) x (3)]);
    };

}

done_testing();
