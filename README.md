# xls-viewer


## Overview

CLI-based Excel viewer


## How to use

- `$ node xls-viewer [options] <filepath>`
  - `[options]`
    - `--sheets=sheetname1,sheetname2,..`
      - specify sheets by name separated by comma(,)
      - default: all sheets
    - `--row_max_width=num`
      - specify maximum width for each rows
      - default: 20
    - `--border=num`
      - specify 0 if border should not be displayed
      - default: 1
    - `--formula=num`
      - specify 1 if formula(not calculated value) should be displayed
      - default: 0
    - `--a1=num`
      - specify 1 if display cells from A1
      - default: 0


## Reference

- https://docs.sheetjs.com/docs/getting-started/installation/nodejs/#legacy-endpoints

- https://qiita.com/Kazunori-Kimura/items/29038632361fba69de5e



## Licensing

This code is licensed under MIT.


## Copyright

2024  [K.Kimura @ Juge.Me](https://github.com/dotnsf) all rights reserved.
