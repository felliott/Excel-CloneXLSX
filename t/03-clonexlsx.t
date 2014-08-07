use strict;
use warnings;

use Test::Fatal;
use Test::More 0.98;

use Excel::CloneXLSX;
use File::Temp qw(tempfile);
use Safe::Isa;
use Spreadsheet::ParseExcel::Utility qw(sheetRef);


my $clone_easy = 't/data/clone-easy.xlsx';
{
    my ($fh, $tempfile) = tempfile();
    Excel::CloneXLSX->new({
        from => $clone_easy,
        to   => $tempfile,
    })->clone;

    compare_xlsx($clone_easy, $tempfile);
}

done_testing();


sub compare_xlsx {
    my ($orig, $clone) = @_;

    my @xlsxs = (
        {filename => $orig},
        {filename => $clone},
    );

    for my $xlsx (@xlsxs) {
        die "Nerp: $xlsx->{filename}" unless (-e $xlsx->{filename});
        open my $fh, '<', $xlsx->{filename} or die "Can't open $xlsx->{filename}: $!";
        $xlsx->{parser} = Excel::CloneXLSX::WrappedParser->new({
            filehandle => $fh,
        });
        $xlsx->{workbook} = $xlsx->{parser}->workbook;

        $xlsx->{wkst_names} = [map {$_->get_name} $xlsx->{workbook}->worksheets];
    }


    for my $wkst_name (@{ $xlsxs[0]->{wkst_names} }) {
        my @wksts = map {{wkst => $_->{workbook}->worksheet($wkst_name)}} @xlsxs;

        for my $wkst (@wksts) {
            @{ $wkst }{ qw(row_min row_max col_min col_max) }
                = ($wkst->{wkst}->row_range(),$wkst->{wkst}->col_range());
        }

        is $wksts[1]->{row_min}, $wksts[0]->{row_min}, 'row_mins are the same';
        is $wksts[1]->{row_max}, $wksts[0]->{row_max}, 'row_maxs are the same';
        is $wksts[1]->{col_min}, $wksts[0]->{col_min}, 'col_mins are the same';
        is $wksts[1]->{col_max}, $wksts[0]->{col_max}, 'col_maxs are the same';

        # for my $row ($wksts[0]->{row_min}..$wksts[0]->{row_max}) {
        #     _cmp_formats(
        #         $xlsxs[1]->{parser}->get_row_format($wkst_name, $row),
        #         $xlsxs[0]->{parser}->get_row_format($wkst_name, $row),
        #         "Row formats for $row are the same",
        #     );
        # }

        # for my $col ($wksts[0]->{col_min}..$wksts[0]->{col_max}) {
        #     _cmp_formats(
        #         $xlsxs[1]->{parser}->get_col_format($wkst_name, $col),
        #         $xlsxs[0]->{parser}->get_col_format($wkst_name, $col),
        #         "Col formats for $col are the same",
        #     );
        # }

        for my $row ($wksts[0]->{row_min}..$wksts[0]->{row_max}) {
            for my $col ($wksts[0]->{col_min}..$wksts[0]->{col_max}) {
                # _cmp_formats(
                #     $xlsxs[1]{parser}->get_formatting_for_cell($wkst_name, $row, $col),
                #     $xlsxs[0]{parser}->get_formatting_for_cell($wkst_name, $row, $col),
                #     "Cell formats for ($row, $col) are equal",
                # );

                is(
                    ($wksts[1]->{wkst}->get_cell($row,$col)->$_call_if_object('value') // ''),
                    ($wksts[0]->{wkst}->get_cell($row,$col)->$_call_if_object('value') // ''),
                    "Cell values for ($row, $col) are equal");
            }
        }


    }


}


sub _cmp_formats {
    my $fmt1 = shift;
    my $fmt2 = shift;

    # my @ignore_fields = qw(
    #     IgnoreFont FmtIdx IgnoreAlignment FontNo IgnoreNumberFormat Hidden
    #     Merged IgnoreFill IgnoreBorder IgnoreProtection Wrap
    # );
    # for my $key (@ignore_fields) {
    #     delete $fmt1->{$key};
    #     delete $fmt2->{$key};
    # }
    # is_deeply $fmt1, $fmt2, @_;

    my @cmp_fields = qw(AlignH AlignV BdrColor BdrDiag BdrStyle Fill);
    my %cmp1 = map {$_ => $fmt1->{$_}} @cmp_fields;
    my %cmp2 = map {$_ => $fmt2->{$_}} @cmp_fields;
    is_deeply \%cmp1, \%cmp2;
}
