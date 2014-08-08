requires 'perl', '5.008001';

requires 'Excel::CloneXLSX::Format';
requires 'Excel::Writer::XLSX';
requires 'IO::File';
requires 'Moo';
requires 'namespace::clean';
requires 'Safe::Isa';
requires 'Spreadsheet::ParseXLSX';
requires 'Types::Standard';
requires 'Type::Tiny';
requires 'XML::Twig';

on 'test' => sub {
    requires 'File::Temp';
    requires 'Module::Pluggable';
    requires 'Spreadsheet::ParseExcel';
    requires 'Test::Fatal';
    requires 'Test::More', '0.98';
};

