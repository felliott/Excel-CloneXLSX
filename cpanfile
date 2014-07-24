requires 'perl', '5.008001';

requires 'Excel::Writer::XLSX';
requires 'Moo';
requires 'namespace::clean';
requires 'Safe::Isa';
requires 'Spreadsheet::ParseXLSX';
requires 'XML::Twig';

on 'test' => sub {
    requires 'Test::More', '0.98';
};

