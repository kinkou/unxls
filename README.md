# Unxls

[![Gem Version](https://badge.fury.io/rb/unxls.svg)](https://badge.fury.io/rb/unxls)
[![Build Status](https://travis-ci.com/kinkou/unxls.svg?branch=master)](https://travis-ci.com/kinkou/unxls)
[![Maintainability](https://api.codeclimate.com/v1/badges/e6cba5968592fa84807d/maintainability)](https://codeclimate.com/github/kinkou/unxls/maintainability)

Unxls is a parser for Microsoft Excel XLS binary file formats with the ultimate goal to cover the entire format specification and all BIFF versions. It is written with these ideas in mind:

- Single responsibility. This gem's only goal is to decode XLS files (as scrupulously as possible). It is quite a complex task already, and combining it with other functionality (like manipulating file structure, generating new XLS files, etc), as other libraries do, is not a good idea, in my opinion.
- Extensibility. It should be relatively easy for other developers to add support for new record types. Parsing records is not very difficult, it just requires some time for reading and understanding the official specification, which is fairly comprehensible.
- Interoperability. The output of the gem is a plain Ruby data structure, not an object, and thus can be easily exported to other formats (like JSON or XML).
- [Productionability](https://en.wiktionary.org/wiki/productionable). One of the goals of the project is to optimize the library for speed and memory reasonably. XLS files can get big, whereas only a fraction of the data might be needed, so a possibility to specify what records from which part of the file to extract is planned for implementation.

## How to use
```bash
gem install unxls
```

and then

```ruby
require 'unxls'

Unxls.parse('file.xls')
```

### Resulting data structure
```ruby
{
  workbook_stream: [
    # Global substream is always at index 0
    {
      # Single-type records, like BOF, are saved as a hash value
      BOF: {
        vers: 1536, # original value
        vers_d: :BIFF8, # decoded value
        # … other fields
      },

      # … other records

      # Multiple-type records, like Font, are saved into an array
      # in the same order as encountered in the original file
      Font: [
        {
          bls: 400, # original value
          bls_d: :BLSNORMAL, # decoded value
          # … other fields
        },
        # … other Font records
      ]

      # … other records

    },

    # Worksheet substreams follow starting from index 1
    {
      # … substream records
    },

    # Last in the array is the hash with data used for quick cell/hyperlink/etc. value lookup
    {
      # :"<worksheet index>_<row index>_<column index>" => :<record name>_<record index>
      :"1_0_0" => :LabelSst_0,
      # …

      hlinks: {
        # :"<worksheet index>_<row index>_<column_index>" => :<HLink index>
        :"1_1_0" => 0,
        # …
      },

      # Same as :hlinks
      hlinktooltips: {
        # …
      },
      notes: {
        # …
      },

      # Bounding box of the area of non-empty cells (per worksheet), which,
      # unlike the worksheet's own Dimensions record, counts cells
      # that have no value but have formatting, as empty.
      dimensions: {
        1 => { rmin: 0, rmax: 2, cmin: 1, cmax: 3 },
        # … other worksheets' dimensions
      }

    }
  ],

  # other parsed storages or streams

}
```

## What you get is NOT what you see
One important thing that people often miss is that Unxls and other Excel-related Ruby libraries, like [Spreadsheet](https://github.com/zdavatz/spreadsheet) or [rubyXL](https://github.com/weshatheleopard/rubyXL), only decode (in most cases) data from files __as it is__. What you see in Excel when you open the file is a result of Excel's interpretation of the data from the file, and the difference can be significant.
For example, Excel's feature called number formats allow users display the same cell contents in a huge variety of ways. A number ```1800``` can be displayed as, depending on the specified format, as either ```1.8E+03```, ```$1800.00``` or ```1-800-COOKIES```. Number formats is in fact a kind of templating language, which over the years have grown to support [a lot of features](https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68), including conditions, country codes, text coloring and whatnot. However, the cell value itself will be persisted as a 64-bit floating-point number (if ```RK``` or ```MulRk``` type of record is used) or a 30-bit floating-point or signed integer number (if ```Number``` type is used) and a corresponding number format will be saved as a string (```"0,0E+00"``` in the first case), but that's about it. To be able to get exact same results as Excel __displays__, one needs to implement certain functionality of Excel (like [this JS library](https://github.com/SheetJS/ssf) does, for example).

## Coverage
At the moment, the gem decodes only the Globals and Worksheet substreams from Workbook stream of the BIFF8 OLE compound files.

### Globals substream
These records are decoded:  
General: BOF, BoundSheet8, CalcPrecision, CodePage, Country, Date1904, FilePass  
Cell values: SST  
Cell formatting: DXF, Font, Format, Palette, Style, StyleExt, TableStyle, TableStyleElement, TableStyles, Theme, XF, XFExt

### Worksheet substream
These records are decoded:  
General: Dimensions, MergeCells, WsBool  
Cell values: Blank, BoolErr, Formula, HLink, HLinkTooltip, LabelSst, MulBlank, MulRk, Note, Number, RK, String  
Conditional formatting: CF, CF12, CFEx, CondFmt, CondFmt12  
Cell formatting: ColInfo, PhoneticInfo, Row  
Table (former List): Feature11, Feature12, List12  
PivotTable: SxView, SXViewEx9  
Other: Obj (partial support), TxO

## Supported encryption algorithms
Unxls can decrypt password-protected files, if the password is known, or if the default password (```VelvetSweatshop```) is used. It supports all encryption types that can be used in binary Excel files: XOR obfuscation (Method 1), 40-bit RC4 encryption and CryptoAPI encryption. However, I couldn't find any files that would use CryptoAPI type with AES192 or AES256 ciphers, so only AES128 and RC4 were tested to date.  
To provide a password for parsing an encrypted file, specify it in the arguments like this:
```ruby
Unxls.parse('file.xls', { password: 'password' })
```

## Development
The gem is written using as much plain Ruby as possible, with minimum dependencies to just work with future Ruby versions and require zero maintenance.  
The C extension for bitwise operations ```bit_ops.c``` was added for the sake of experimentation and will probably be removed in future versions, because it showed almost no performance benefits over the pure Ruby (2.6) implementation.

### Under construction
It is important to mention that the gem is still under construction and things **may** break in the future versions. A number of decisions regarding its architecture were postponed, because the XLS file specification is very large (over 1000 pages) and there might be records or structures for which parsing has not yet been implemented, which in turn might require serious changes in the parser algorithm. However, most of the gem code that deals with concrete records and structures is written in procedural style and thus can be transferred between versions with little or no changes.  
Another important consideration is that the decoded data is often not very useful as is. One still needs to know a lot to put the right structures together and interpret the result to get meaningful output. There is a need for a convenient interface to retrieve things like for example values or border styles of concrete cells, and eventually it will be implemented as a separate gem. The functionality is however needed for testing, so it is being accumulated and experimented with in the ```Unxls::Biff8::Browser``` class.

### Specifications
If you'd like to find out what the fields in the resulting parsed structure mean and how to put them together, or if you'd like to contribute, please read the specification.  
The best place to start in my opinion is [Excel Binary File Format (.xls) Structure](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cd03cb5f-ca02-4934-a391-bb674cb8aa06) (in case the link breaks, try googling ```[MS-XLS]```), the latest version of the specification, maintained by Microsoft itself. This and other documents it refers to, like ```[MS-DTYP]```, ```[MS-OFFCRYPTO]```, ```[MS-OLEPS]```, ```[MS-OSHARED]```, ```[MS-UCODEREF]``` describe in detail all of the BIFF8 internals.  
There are two more documents that have information about older BIFF versions and feature valuable details the official documentation lacks:
 - [Microsoft Office Excel 97-2007 - Binary File Format Specification](http://download.microsoft.com/download/5/0/1/501ED102-E53F-4CE0-AA6B-B0F93629DDC6/Office/Excel97-2007BinaryFileFormat(xls)Specification.pdf) (in case the link breaks, try googling ```Excel97-2007BinaryFileFormat(xls)Specification.pdf```) and
 - [OpenOffice.org's Documentation of the Microsoft Excel File Format](https://www.openoffice.org/sc/excelfileformat.pdf) (in case the link breaks, try googling ```excelfileformat.pdf```).

## Debugging utils
These utils in the ```bin``` folder are included to make testing and debugging easier:
- ```unxls [options] [file]``` parses a file or a directory of files. The former is useful for quickly displaying the contents of a file, and the latter for mass smoke tests of existing or new record/structure parsing methods. To learn more, see ```unxls -h```.
- ```console``` starts a Pry session, with Unxls ```require```d.
- ```records8 [file]``` (for BIFF8 files), ```records5 [file]``` (for BIFF5 files), ```records2``` (for BIFF2 files) simply display, in succession, what records the parsed substreams contain.

## Requirements
Ruby >= 2.1. Enjoy!
