# NAME

Excel::CloneXLSX - Clone an XLSX file and add new rows

# SYNOPSIS

    use Excel::CloneXLSX;

    # copy old.xlsx to new.xlsx
    #  (like a worse version of cp)
    Excel::CloneXLSX->new({
      from => 'old.xlsx', to  => 'new.xlsx',
    })->clone;

    # copy old.xlsx to new.xlsx, but just 'Sheet 2'
    Excel::CloneXLSX->new({
      from => 'old.xlsx', to => 'new.xlsx',
      worksheets => ['Sheet 2'],
    })->clone;

    # copy old.xlsx to new.xlsx, but just 'Sheet 2'
    Excel::CloneXLSX->new({
      from => 'old.xlsx', to => 'new.xlsx',
      worksheets => ['Sheet 1'],
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

# DESCRIPTION

Excel::CloneXLSX is a module for cloning an Excel file while being
able to insert new rows.  It's not very smart.  It will iterate
through the rows of the old spreadsheet, copying them to the new
spreadsheet, occasionally adding new rows according to spec.

# ATTRIBUTES

## from

**Required**.  The spreadsheet to clone.  Can be passed as a filename,
a filehandle, or a reference to a scalar.

## to

**Required**.  The output spreadsheet.  Can be passed as a filename, a
filehandle, or reference to a scalar.  If you give it an existing
spreadsheet, that spreadsheet will be overwritten.  This module does
not modify existing spreadsheets, it creates new one with possible
insertions.

If you coerce from a filehandle, you will need to call
`seek $fh, 0, 0;` on the handle to reset it for reading.

## worksheet\_names

A list of worksheet namess to restrict the cloning to.  If you do not
set this, all worksheets in the ["from"](#from) workbook will be cloned.

## new\_rows

A specification for new rows to insert.

# METHODS

## clone

It... clones, imperfectly.

This method returns nothing, but calls `->close()` on the ["to"](#to)
spreadsheet.  Once this is done, the spreadsheet will be ready for
reading.

# LICENSE

Copyright (C) Fitz Elliott.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

# AUTHOR

Fitz Elliott <felliott@fiskur.org>
