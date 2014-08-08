use strict;
use warnings;

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

    my ($row_min, $row_max) = $wkst->row_range();
    my ($col_min, $col_max) = $wkst->col_range();

    subtest 'Cell Ranges' => sub {
        is $row_min, 0, 'Rows start at 0';
        is $row_max, 4, 'Rows end at 4';
        is $col_min, 0, 'Cols start at 0';
        is $col_max, 2, 'Cols end at 2';
    };


    subtest 'Cell Contents' => sub {
        my $contents;
        for my $row ($row_min..$row_max) {
            for my $col ($col_min..$col_max) {
                $contents->[$row][$col] = $wkst->get_cell($row, $col)->value;
            }
        }

        my @expect = (
            [qw(dog foo car)],
            [qw(moo boo foo)],
            [qw(bar car par)],
            [qw(aaa bbb ccc)],
            [qw(quux ducks trucks)],
        );
        is_deeply $contents, \@expect, 'Contents are as expected';
    };

    subtest 'Cell Formats' => sub {
        my $bgcolors;
        for my $row ($row_min..$row_max) {
            for my $col ($col_min..$col_max) {
                my $fmt = $parser->get_cell_format('Append', $row, $col);
                $bgcolors->[$row][$col] = $fmt && $fmt->{Fill}
                    ? lc($fmt->{Fill}[1]) : undef;
            }
        }

        my @expect = (
            [undef, undef, undef,],
            ['#ff0000', '#ff0000', '#ff0000',],
            [undef, undef, undef,],
            ['#ff0000', '#0000ff', '#ff0000',],
            ['#0000ff', '#0000ff', '#0000ff',],
        );
        is_deeply $bgcolors, \@expect, 'Formatting is as expected';
    };
}

done_testing();

