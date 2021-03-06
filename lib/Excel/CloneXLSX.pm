package Excel::CloneXLSX;

use Excel::CloneXLSX::Format qw(translate_xlsx_format);
use Excel::CloneXLSX::Types qw(CloneXlsxParser CloneXlsxWriter);
use Safe::Isa;
use Types::Standard -types;

use Moo;
use namespace::clean;

our $VERSION = "0.03";


has from => (
    is       => 'ro',
    isa      => CloneXlsxParser,
    coerce   => CloneXlsxParser->coercion,
    required => 1,
);
has to => (
    is       => 'ro',
    isa      => CloneXlsxWriter,
    coerce   => CloneXlsxWriter->coercion,
    required => 1,
);

has worksheet_names => (
    is      => 'ro',
    isa     => ArrayRef[Str],
    default => sub { [] },
);

has _worksheets => (
    is      => 'ro',
    isa     => ArrayRef[InstanceOf['Spreadsheet::ParseExcel::Worksheet']],
    lazy    => 1,
    builder => 1,
);
sub _build__worksheets {
    my ($self) = @_;
    my @names = @{$self->worksheet_names};
    return @names ? [map {$self->from->worksheet($_)} @names]
        : [$self->from->worksheets()];
}

has new_rows => ( is => 'ro', default => sub { {} }, );


sub clone {
    my ($self) = @_;

    for my $old_wkst (@{ $self->_worksheets }) {
        my $new_wkst     = $self->to->add_worksheet($old_wkst->get_name);
        my $old_tabcolor = $old_wkst->get_tab_color();
        $new_wkst->set_tab_color($old_tabcolor) if ($old_tabcolor);

        my ($row_min, $row_max) = $old_wkst->row_range();
        my ($col_min, $col_max) = $old_wkst->col_range();

        my @fmts = map {
            $self->to->add_format( %{ translate_xlsx_format($_) } )
        } @{ $self->from->workbook->{Format} };

        my $row_heights = $old_wkst->get_row_heights();
        my $col_widths  = $old_wkst->get_col_widths();
        for my $col ($col_min..$col_max) {
            my $col_fmt_no = $old_wkst->{ColFmtNo}[$col];
            $new_wkst->set_column(
                $col, $col, $col_widths->[$col],
                (defined $col_fmt_no ? $fmts[ $col_fmt_no ] : undef),
            );
        }

        my $row_offset = 0;
        $row_offset += $self->_append_rows_after(
            $old_wkst, $new_wkst, -1, $row_offset
        );

        for my $row ($row_min..$row_max) {
            my $row_fmt_no = $self->from->row_format_no->{$old_wkst->get_name}[$row];
            $new_wkst->set_row(
                $row+$row_offset, $row_heights->[$row],
                (defined $row_fmt_no ? $fmts[ $row_fmt_no ] : undef),
            );

            for my $col ($col_min..$col_max) {
                my $cell    = $old_wkst->get_cell($row, $col);
                my $content = $cell->$_call_if_object('unformatted');
                my $old_fmt = $self->from->get_cell_format(
                    $old_wkst->{Name}, $row, $col
                );
                my $new_format = $old_fmt
                    ? $self->to->add_format(%{ translate_xlsx_format($old_fmt) })
                        : undef;

                $new_wkst->write($row+$row_offset, $col, $content, $new_format);
            }

            $row_offset += $self->_append_rows_after(
                $old_wkst, $new_wkst, $row, $row_offset
            );
        }
    }

    $self->to->close;
    return;
}


sub _append_rows_after {
    my ($self, $old_wkst, $new_wkst, $row, $row_offset) = @_;

    return 0 unless (exists $self->new_rows->{$row});

    my $rows_added = 0;
    my ($col_min, $col_max) = $old_wkst->col_range();
    my $new_rows = $self->new_rows->{$row};
    for my $new_row (@$new_rows) {
        $rows_added++;

        for my $col ($col_min..$col_max) {
            my $new_cell = $new_row->[$col];

            my ($new_content, $new_format)
                = @{$new_cell}{qw(content format)};

            if (ref $new_content eq 'ARRAY') {
                my ($delta_row, $delta_col) = @$new_content;
                $new_content = $old_wkst->get_cell(
                    $row+$delta_row,
                    $col+$delta_col,
                )->$_call_if_object('unformatted') || undef;
            }

            if (ref $new_format eq 'ARRAY') {
                my ($delta_row, $delta_col) = @$new_format;
                my $old_fmt = $self->from->get_cell_format(
                    $old_wkst->{Name},
                    $row+$delta_row,
                    $col+$delta_col,
                );

                $new_format = $old_fmt
                    ? $self->to->add_format(%{ translate_xlsx_format($old_fmt) })
                        : undef;
            }
            elsif (ref $new_format eq 'HASH') {
                $new_format = $self->to->add_format( %$new_format );
            }

            my @args = ( $row+$row_offset+$rows_added, $col, $new_content );
            push @args, $new_format if ($new_format);
            $new_wkst->write(@args);
        }
    }

    return $rows_added;
}

1;
__END__

=encoding utf-8

=head1 NAME

Excel::CloneXLSX - Clone an XLSX file and add new rows

=head1 SYNOPSIS

    use Excel::CloneXLSX;

    # copy old.xlsx to new.xlsx
    #  (like a worse version of cp)
    Excel::CloneXLSX->new({
      from => 'old.xlsx', to  => 'new.xlsx',
    })->clone;

    # copy old.xlsx to new.xlsx, but just 'Sheet 2'
    Excel::CloneXLSX->new({
      from => 'old.xlsx', to => 'new.xlsx',
      worksheet_names => ['Sheet 2'],
    })->clone;

    # copy old.xlsx to new.xlsx, but just 'Sheet 2'
    Excel::CloneXLSX->new({
      from => 'old.xlsx', to => 'new.xlsx',
      worksheet_names => ['Sheet 1'],
      new_rows => {
        '0' => [
          [ # this will be second row in the new worksheet
            {content => '', format => ''},
            {content => '', format => ''},
            {content => '', format => ''},
            {content => '', format => ''},
          ],
        ],
    })->clone;

=head1 DESCRIPTION

Excel::CloneXLSX is a module for cloning an Excel file while being
able to insert new rows.  It's not very smart.  It will iterate
through the rows of the old spreadsheet, copying them to the new
spreadsheet, occasionally adding new rows according to spec.


=head1 ATTRIBUTES

=head2 from

B<Required>.  The spreadsheet to clone.  Can be passed as a filename,
a filehandle, or a reference to a scalar.

=head2 to

B<Required>.  The output spreadsheet.  Can be passed as a filename, a
filehandle, or reference to a scalar.  If you give it an existing
spreadsheet, that spreadsheet will be overwritten.  This module does
not modify existing spreadsheets, it only creates new ones.

If you coerce from a filehandle, you will need to call
C<< seek $fh, 0, 0; >> on the handle to reset it for reading.

=head2 worksheet_names

A list of worksheet names to clone.  If you do not set this, all
worksheets in the L</from> workbook will be cloned.

=head2 new_rows

A specification for new rows to insert.

=head1 METHODS

=head2 clone

It... clones, imperfectly.

This method returns nothing, but calls C<< ->close() >> on the L</to>
spreadsheet.  Once this is done, the spreadsheet will be ready for
reading.  If you passed a filehandle as the C<to()> argument, you will
need to call C<< seek $fh, 0, 0; >> to reset it for reading.


=head1 LICENSE

Copyright (C) Fitz Elliott.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.


=head1 AUTHOR

Fitz Elliott E<lt>felliott@fiskur.orgE<gt>

=cut
