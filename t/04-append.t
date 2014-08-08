use strict;
use warnings;

use lib q{t/lib};

use t::CloneXLSX::Utils qw(:all);
use Test::Fatal;
use Test::More 0.98;

use Excel::CloneXLSX;
use File::Temp qw(tempfile);


my $append = 't/data/append.xlsx';
{
    my ($fh, $tempfile) = tempfile();
    Excel::CloneXLSX->new({
        from => $append,
        to   => $tempfile,
        new_rows => {
            '-1' => [
                [
                    {content => 'dog',  format => undef,},
                    {content => [1,1],  format => undef,},
                    {content => [2,-1], format => undef,},
                ],
            ],
            '1' => [
                [
                    {content => 'aaa', format => [-1,0],},
                    {content => 'bbb', format => [+1,0],},
                    {content => 'ccc', format => [-1,0],},
                ],
            ],
        },
    })->clone;

    open my $new_fh, '<', $tempfile or die "Can't open $tempfile: $!";
    my $parser = Excel::CloneXLSX::WrappedParser->new({
        filehandle => $new_fh,
    });
    my $workbook = $parser->workbook;

    is_deeply [map {$_->get_name} $workbook->worksheets()], [qw(Append)],
        'same workbooks';
    my $wkst = $workbook->worksheet('Append');
    worksheet_range_is($wkst, [0,4], [0,2]);
    cell_contents_are($wkst, [
        [qw(dog  foo   car   )],
        [qw(moo  boo   foo   )],
        [qw(bar  car   par   )],
        [qw(aaa  bbb   ccc   )],
        [qw(quux ducks trucks)],
    ]);
    my ($red, $green, $blue) = ('#ff0000', '#008000', '#0000ff');
    cell_bgcolors_are($parser, $wkst, [
        [undef, undef, undef, ],
        [$red,  $red,  $red,  ],
        [undef, undef, undef, ],
        [$red,  $blue, $red,  ],
        [$blue, $blue, $blue, ],
    ]);
}

done_testing();

