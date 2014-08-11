package Excel::CloneXLSX::Types;

use strict;
use warnings;

use Excel::CloneXLSX::WrappedParser;
use Excel::Writer::XLSX;
use IO::File;
use Type::Library
    -base,
    -declare => qw(
        CloneXlsxInfile  CloneXlsxParser
        CloneXlsxOutfile CloneXlsxWriter
    );
use Type::Utils -all;
use Types::Standard -types;

our $VERSION = "0.03";


declare CloneXlsxInfile, as FileHandle;
coerce CloneXlsxInfile,
    from Str,       via { IO::File->new($_, 'r') },
    from ScalarRef, via { open my $fh, '<', $_; $fh };

class_type CloneXlsxParser, { class => 'Excel::CloneXLSX::WrappedParser' };
coerce CloneXlsxParser,
    from CloneXlsxInfile->coercibles, via {
        my $tmp = to_CloneXlsxInfile($_);
        Excel::CloneXLSX::WrappedParser->new({ filehandle => $tmp })
    };


declare CloneXlsxOutfile, as FileHandle;
coerce CloneXlsxOutfile,
    from Str,       via { IO::File->new($_, '+>') },
    from ScalarRef, via { open my $fh, '+>', $_; $fh };

class_type CloneXlsxWriter, { class => 'Excel::Writer::XLSX' };
coerce CloneXlsxWriter,
    from CloneXlsxOutfile->coercibles, via {
        my $tmp = to_CloneXlsxOutfile($_);
        Excel::Writer::XLSX->new( $tmp )
    };


1;
__END__


=encoding utf-8

=head1 NAME

Excel::CloneXLSX::Types - Type library for Excel::CloneXLSX


=head1 SYNOPSIS

    use Excel::CloneXLSX::Types qw(CloneXlsxWriter);

    my $writer = to_CloneXlsxWriter("myfile.xlsx");
    # or
    my $new_xlsx;
    my $writer = to_CloneXlsxWriter(\$new_xlsx);
    # or
    my $fh = tempfile();
    my $writer = to_CloneXlsxWriter($fh);


=head1 DESCRIPTION

Excel::CloneXLSX::Types defines the types used for argument validation
and coercion by the Excel::CloneXLSX suite of modules.


=head1 EXPORTED TYPES

=head2 CloneXlsxInfile

A seekable filehandle.  Coerces from C<Str> by assuming it's a
filename.  Coerces from C<ScalarRef> by opening a filehandle to the
scalar.

=head2 CloneXlsxParser

An instance of L<Excel::CloneXLSX::WrappedParser>. Uses the coercions
from L</CloneXlsxInfile>.

=head2 CloneXlsxOutfile

A writable filehandle.  Coerces from C<Str> by assuming it's a
filename.  Coerces from C<ScalarRef> by opening a filehandle to the
scalar.

=head2 CloneXlsxWriter

An instance of L<Excel::Writer::XLSX>. Uses the coercions
from L</CloneXlsxOutfile>.


=head1 LICENSE

Copyright (C) Fitz Elliott.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.


=head1 AUTHOR

Fitz Elliott E<lt>felliott@fiskur.orgE<gt>

=cut

