package Excel::CloneXLSX;

use strict;
use warnings;

use Excel::CloneXLSX::WrappedParser;
use Excel::CloneXLSX::Format qw(translate_xlsx_format);
use Excel::Writer::XLSX;
use Safe::Isa;

use Moo;
use namespace::clean;


our $VERSION = "0.01";

has from        => (is => 'ro', lazy => 1, builder => 1);
has from_fh     => (is => 'ro', required => 1);
has from_parser => (is => 'ro', lazy => 1, builder => 1);
sub _build_from { $_[0]->from_parser->parse( $_[0]->from_fh ) }
sub _build_from_parser { Excel::CloneXLSX::WrappedParser->new }


has to    => (is => 'ro', lazy => 1, builder => 1);
has to_fh => (is => 'ro', lazy => 1, builder => 1);
sub _build_to { Excel::Writer::XLSX->new( $_[0]->to_fh ) }
sub _build_to_fh {
    open my $fh, '>', \my $clone or die 'Nope';
    return $fh;
}


has worksheets => (is => 'ro', lazy => 1, builder => 1,);
sub _build_worksheets { [] }

has row_offset => (is => 'ro', default => 0,);
sub incr_row_offset { $_[0]->{row_offset}++ }


has new_rows => (is => 'ro', default => sub { {} });
sub insert_rows_after {
    my ($self, $row) = @_;
    return $self->new_rows->{$row} || undef;
}



sub clone {
    my ($self) = @_;


    for my $old_wkst (map {$self->from->worksheet($_)} @{ $self->worksheets }) {
        my $new_wkst     = $self->to->add_worksheet($old_wkst->get_name);
        my $old_tabcolor = $old_wkst->get_tab_color();
        $new_wkst->set_tab_color($old_tabcolor) if ($old_tabcolor);

        my ($row_min, $row_max) = $old_wkst->row_range();
        my ($col_min, $col_max) = $old_wkst->col_range();

        my $row_heights = $old_wkst->get_row_heights();
        my $col_widths  = $old_wkst->get_col_widths();
        my @col_fmts = map {
            $self->to->add_format( %{ translate_xlsx_format($_) } )
        } @{ $self->from->{Format} };
        for my $col ($col_min..$col_max) {
            $new_wkst->set_column(
                $col, $col, $col_widths->[$col],
                $col_fmts[ $old_wkst->{ColFmtNo}[$col] ],
            );
        }

        my $row_offset = 0;

        for my $row ($row_min..$row_max) {
            $new_wkst->set_row($row+$row_offset, $row_heights->[$row]);

            for my $col ($col_min..$col_max) {
                my $cell       = $old_wkst->get_cell($row, $col);
                my $old_fmt    = $self->from_parser->get_formatting_for_cell(
                    $old_wkst->{Name}, $row, $col
                );
                my $new_format = $old_fmt
                    ? $self->to->add_format(%{ translate_xlsx_format($old_fmt) })
                        : undef;
                $new_wkst->write(
                    $row+$row_offset, $col,
                    ($cell->$_call_if_object('unformatted') || undef),
                    $new_format
                );
            }

            if (my $new_rows = $self->insert_rows_after($row)) {
                for my $new_row (@$new_rows) {
                    $row_offset++;
                    for my $col ($col_min..$col_max) {
                        my $new_cell = $new_row->[$col];

                        my ($new_content, $new_format)
                            = @{$new_cell}{qw(content format)};

                        if (ref $new_content eq 'ARRAY') {
                            my ($delta_row, $delta_col) = @$new_content;
                            $new_content = $old_wkst->get_cell(
                                $row+$delta_row,
                                $col+$delta_row,
                            )->$_call_if_object('unformatted') || undef;
                        }

                        if (ref $new_format eq 'ARRAY') {
                            my ($delta_row, $delta_col) = @$new_format;
                            my $old_fmt = $self->from_parser->get_formatting_for_cell(
                                $old_wkst->{Name},
                                $row+$delta_row,
                                $col+$delta_row,
                            );

                            $new_format = $old_fmt
                                ? $self->to->add_format(%{ translate_xlsx_format($old_fmt) })
                                    : undef;
                        }
                        elsif (ref $new_format eq 'HASH') {
                            $new_format = $self->to->add_format( %$new_format );
                        }

                        my @args = ( $row+$row_offset, $col, $new_content );
                        push @args, $new_format if ($new_format);
                        $new_wkst->write(@args);
                    }
                }
            }
        }
    }

    $self->to->close;
    return $self->to_fh;
}



1;
__END__

=encoding utf-8

=head1 NAME

Excel::CloneXLSX - Clone an XLSX file and add new rows

=head1 SYNOPSIS

    use Excel::CloneXLSX;

    Excel::CloneXLSX->new({
      from => 'old.xlsx',
      new  => 'new.xlsx',
    })->clone;

=head1 DESCRIPTION

Excel::CloneXLSX is a module for cloning an Excel file while being
able to insert new rows.

=head1 LICENSE

Copyright (C) Fitz Elliott.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=head1 AUTHOR

Fitz Elliott E<lt>felliott@fiskur.orgE<gt>

=cut

