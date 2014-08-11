package t::CloneXLSX::Utils;

use strict;
use warnings;

use Test::More 0.98;

use Safe::Isa;

use Exporter 'import';
our %EXPORT_TAGS = (all => [qw(
    worksheet_range_is cell_contents_are cell_bgcolors_are
)]);
Exporter::export_ok_tags('all');


sub worksheet_range_is {
    my ($wkst, $row, $col) = @_;
    subtest 'Cell Ranges' => sub {
        my @ranges = ($wkst->row_range(), $wkst->col_range());
        is $ranges[0], $row->[0], "Rows start at $row->[0]";
        is $ranges[1], $row->[1], "Rows end at $row->[1]";
        is $ranges[2], $col->[0], "Columns start at $col->[0]";
        is $ranges[3], $col->[1], "Columns end at $col->[1]";
    };
}


sub cell_contents_are {
    my ($wkst, $expect) = @_;
    subtest 'Cell Contents' => sub {
        my ($row_min, $row_max) = $wkst->row_range();
        my ($col_min, $col_max) = $wkst->col_range();
        my $contents;
        for my $row ($row_min..$row_max) {
            for my $col ($col_min..$col_max) {
                $contents->[$row][$col]
                    = $wkst->get_cell($row, $col)->$_call_if_object('value');
            }
        }

        is_deeply $contents, $expect, 'Contents are as expected';
    };
}

sub cell_bgcolors_are {
    my ($parser, $wkst, $expect) = @_;
    subtest 'Cell Formats' => sub {
        my ($row_min, $row_max) = $wkst->row_range();
        my ($col_min, $col_max) = $wkst->col_range();
        my $bgcolors;
        for my $row ($row_min..$row_max) {
            for my $col ($col_min..$col_max) {
                my $fmt = $parser->get_computed_cell_format($wkst->get_name(), $row, $col);
                $bgcolors->[$row][$col] = $fmt && $fmt->{Fill}
                    ? lc($fmt->{Fill}[1]) : undef;
            }
        }
        is_deeply $bgcolors, $expect, 'Formatting is as expected';
    };
}


1;
__END__
