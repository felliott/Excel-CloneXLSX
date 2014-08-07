package Excel::CloneXLSX::WrappedParser;

use Types::Standard -types;
use XML::Twig;

use Moo;
extends 'Spreadsheet::ParseXLSX';
use namespace::clean;


has filehandle => (is => 'ro', isa => FileHandle, required => 1);

has workbook => (
    is      => 'ro',
    isa     => InstanceOf['Spreadsheet::ParseExcel::Workbook'],
    lazy    => 1,
    builder => 1,
    handles => [qw(worksheet worksheets)],
);
sub _build_workbook { $_[0]->parse( $_[0]->filehandle ) };

has row_format_no => (
    is      => 'ro',
    isa     => HashRef[ArrayRef[Maybe[Int]]],
    default => sub { {} },
);

has col_format_no => (
    is      => 'ro',
    isa     => HashRef[ArrayRef[Maybe[Int]]],
    lazy    => 1,
    builder => 1,
);
sub _build_col_format_no {
    return {
        map { $_->get_name() => [ @{$_->{ColFmtNo}} ] } $_[0]->worksheets
    };
}

has cell_formats => (
    is      => 'ro',
    isa     => HashRef[ArrayRef[ArrayRef]],
    default => sub { {} },
);


around _parse_sheet => sub {
    my ($orig, $self, $sheet, $sheet_file) = @_;
    $self->$orig($sheet, $sheet_file);

    my @formats;

    my $sheet_xml = XML::Twig->new(
        twig_roots => {
            'sheetData/row' => sub { # get default row format ids
                my ($twig, $row) = @_;
                $self->row_format_no->{$sheet->{Name}}[ $row->att('r') - 1 ]
                    =  ($row->att('s') && $row->att('customFormat'))
                        ? $row->att('s') : undef;
                $twig->purge;
            },
            'sheetData/row/c' => sub { # get cell-specific format ids
                my ($twig, $cell) = @_;
                my ($row, $col) = $self->_cell_to_row_col($cell->att('r'));
                $formats[$row][$col] = defined($cell->att('s'))
                    ? $sheet->{_Book}{Format}[$cell->att('s')]
                        : undef;
            },
        }
    );
    $sheet_xml->parse( $sheet_file );

    my ($row_min, $row_max, $col_min, $col_max)
        = @{$sheet}{qw(MinRow MaxRow MinCol MaxCol)};

    # set format for each cell:
    #  cell-specific > row-default > col-default
    for my $row ($row_min..$row_max) {
        for my $col ($col_min..$col_max) {
            next if (defined $formats[$row][$col]);

            # check for row format
            if (defined $self->row_format_no->{$sheet->{Name}}[$row]) {
                $formats[$row][$col] = $sheet->{_Book}{Format}[$self->row_format_no->{$sheet->{Name}}[$row]];
            }
            elsif (defined $sheet->{ColFmtNo}[$col]) {
                $formats[$row][$col] = $sheet->{_Book}{Format}[$sheet->{ColFmtNo}[$col]];
            }
        }
    }

    $self->cell_formats->{$sheet->{Name}} = \@formats;
};


sub get_formatting_for_cell {
    my ($self, $sheet_name, $row, $col) = @_;
    return $self->cell_formats->{$sheet_name}[$row][$col];
}



1;
__END__

=encoding utf-8

=head1 NAME

Excel::CloneXLSX::WrappedParser - Wrapper for Spreadsheet::ParseXLSX


=head1 SYNOPSIS

    use Excel::CloneXLSX::WrappedParser;

    my $parser = Excel::CloneXLSX::WrappedParser->new;
    my $workbook = $parser->parse('foo.xlsx');

    # get format for cell G5 (row 4, col 6)
    my $format = $parser->get_formatting_for_cell(
      'Sheet 1', 4, 6
    );


=head1 DESCRIPTION

Excel::CloneXLSX::WrappedParser wraps the L<Spreadsheet::ParseXLSX>
module into order to hook into its XML parsing and save additional
information we need for L<Excel::CloneXLSX>.

The extra information we currently save is the computed formats for
every cell in the workbook.  L<Spreadsheet::ParseExcel> doesn't provide
a way to get at the formatting for a cell with no defined content.


=head1 ATTRIBUTES

=head2 filehandle

B<Required>.  A file handle for the spreadsheet to be parsed.

=head2 workbook

The L<Spreadsheet::ParseExcel::Workbook> object returned by
L<Spreadsheet::ParseXLSX>'s C<parse()> method.

=head2 row_format_no

A hashref of arrayrefs of default row format ids for each worksheet.
L<Spreadsheet::ParseExcel::Worksheet> objects store all the format
objects for a worksheet in the C<< $worksheet->{Format} >> arrayref.
The default column format ids are stored in an internal key called
C<< $worksheet->{ColFmtNo} >>.  For instance, to get the default column format
for column C, you would do:
C<< $worksheet->{Format}[ $worksheet->{ColFmtNo}[2] ] >>. The
C<row_format_no> attribute provides the same for default row formats.
Since we have to store properties for the entire workbook, the row
formats are indexed by worksheet name and row number.  E.g. the
default format for row 7 of sheet 'Sheet 1' is available through
 C<< $worksheet->{Format}[ $parser->row_format_no->{'Sheet 1'}[6] ] >>.

=head2 col_format_no

For symmetry, we provide a C<col_format_no> method that is a hashref
of C<< $sheet_name => \@col_formats >>.

=head2 cell_formats

A hashref of 2-D arrayrefs that stores the
L<Spreadsheet::ParseExcel::Format> objects for each cell in each
worksheet.  Get at them via:
C<< $parser->cell_formats->{$sheet_name}[$row][$col] >>, or just use the
C<get_formatting_for_cell()> method.

A cell can have three types of formats applied to it: cell-specific,
row-default, column-default.  A cell-specific format will take
precedence over a row-specfic format, which takes precedence over a
column-specific format.  They are not additive.


=head1 METHODS

=head2 worksheet

Delegated method to L</workbook>'s C<worksheet> method.

=head2 worksheets

Delegated method to L</workbook>'s C<worksheets> method.

=head2 get_formatting_for_cell( $sheet_name, $row, $col )

Returns the L<Spreadsheet::ParseExcel::Format> object for the cell at
(C<$row>, C<$col>) in sheet C<$sheet_name>.  C<$row> and C<$col> are
the 0-based coordinates of the cell. E.g: "A5" = (4,0), "C7" = (6,2)
The C<sheetRef()> method of L<Spreadsheet::ParseExcel::Utility> will
convert Excel notation to zero-indexing.


=head1 LICENSE

Copyright (C) Fitz Elliott.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.


=head1 AUTHOR

Fitz Elliott E<lt>felliott@fiskur.orgE<gt>

=cut

