package Excel::CloneXLSX::WrappedParser;

use XML::Twig;

use Moo;
extends 'Spreadsheet::ParseXLSX';
use namespace::clean;


has _files => (is => 'rw');

has _cell_formats => (is => 'ro', default => sub { {} },);
has row_format_no => (is => 'ro', default => sub { [] },);

around _extract_files => sub {
    my $orig = shift;
    my $self = shift;
    my $files = $self->$orig(@_);
    $self->_files($files);
    return $files;
};


around _parse_sheet => sub {
    my ($orig, $self, $sheet, $sheet_file) = @_;
    $self->$orig($sheet, $sheet_file);

    my @formats;

    my $sheet_xml = XML::Twig->new(
        twig_roots => {
            'sheetData/row' => sub {
                my ($twig, $row) = @_;
                $self->row_format_no->[ $row->att('r') - 1 ]
                    =  ($row->att('s') && $row->att('customFormat'))
                        ? $row->att('s') : undef;
                $twig->purge;
            },
            'sheetData/row/c' => sub {
                my ($twig, $cell) = @_;
                my ($row, $col) = $self->_cell_to_row_col($cell->att('r'));
                $formats[$row][$col] = $sheet->{_Book}{Format}[$cell->att('s')];

            },
        }
    );
    $sheet_xml->parse( $sheet_file );

    my ($row_min, $row_max, $col_min, $col_max)
        = @{$sheet}{qw(MinRow MaxRow MinCol MaxCol)};

    for my $row ($row_min..$row_max) {
        for my $col ($col_min..$col_max) {
            next if (defined $formats[$row][$col]);

            # check for row format
            if (defined $self->row_format_no->[$row]) {
                $formats[$row][$col] = $sheet->{_Book}{Format}[$self->row_format_no->[$row]];
            }
            elsif (defined $sheet->{ColFmtNo}[$col]) {
                $formats[$row][$col] = $sheet->{_Book}{Format}[$sheet->{ColFmtNo}[$col]];
            }
        }
    }

    $self->_cell_formats->{$sheet->{Name}} = \@formats;
};


sub get_formatting_for_cell {
    my ($self, $sheet_name, $row, $col) = @_;
    return $self->_cell_formats->{$sheet_name}[$row][$col];
}

1;
