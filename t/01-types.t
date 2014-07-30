use strict;
use warnings;

use Test::More 0.98;

use Excel::CloneXLSX::Types qw(to_CloneXlsxParser to_CloneXlsxWriter);
use File::Temp qw(tempfile);

my $basic = 't/data/basic.xlsx';

{
    ok my $parser = to_CloneXlsxParser($basic);
    isa_ok $parser, 'Excel::CloneXLSX::WrappedParser';
}

{
    open my $fh, '<', $basic or die "Can't open $basic: $!";
    ok my $parser = to_CloneXlsxParser($fh);
    isa_ok $parser, 'Excel::CloneXLSX::WrappedParser';
    close $fh;
}

{
    open my $fh, '<', $basic or die "Can't open $basic: $!";
    binmode($fh);
    local $/ = undef;
    my $xlsx = <$fh>;
    ok my $parser = to_CloneXlsxParser(\$xlsx);
    isa_ok $parser, 'Excel::CloneXLSX::WrappedParser';
    close $fh;
}

{
    my $scalar;
    my @outputs = (
        [(tempfile())[1], 'filename',  ],
        [(tempfile())[0], 'filehandle',],
        [\$scalar,        'scalarref', ],
    );

    for my $output (@outputs) {
        my ($out, $type) = @$output;
        ok my $writer = to_CloneXlsxWriter($out), "can coerce $type to writer";
        isa_ok $writer, 'Excel::Writer::XLSX';
        ok $writer->close, "  ...and closeable, no less";
    }
}


done_testing;
