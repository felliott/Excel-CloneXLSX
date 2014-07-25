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
    from Str,       via { IO::File->new($_, 'rw') },
    from ScalarRef, via { open my $fh, '+>', $_; $fh };

class_type CloneXlsxWriter, { class => 'Excel::Writer::XLSX' };
coerce CloneXlsxWriter,
    from CloneXlsxOutfile->coercibles, via {
        my $tmp = to_CloneXlsxOutfile($_);
        Excel::Writer::XLSX->new( $tmp )
    };


1;
__END__
