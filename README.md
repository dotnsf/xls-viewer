# xls-viewer


## Overview

CLI-based Excel viewer


## How to install

- Just copy executables under `releases/` folder into executable folders(ex. `/usr/local/bin`)
  - `releases/xls-viewer`
    - for linux-amd64
  - `releases/xls-viewer.exe`
    - for windows-amd64


## How to use

- `$ node xls-viewer.js [options] <filepath>`, or
- `$ xls-viewer [options] <filepath>`

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
    - `--label=num`
      - specify 1 if display labels of row/column
      - default: 0

  
## How to build executable binaries

- Install **pkg**
  - `$ sudo npm install -g pkg`

- Build binary for linux/amd64
  - `$ pkg -t node16-linux-x64 -o xls-viewer xls-viewer.js`

- Build binary for windows/amd64
  - `$ pkg -t node16-win-x64 -o xls-viewer.exe xls-viewer.js`


## Reference

- https://docs.sheetjs.com/docs/getting-started/installation/nodejs/#legacy-endpoints

- https://qiita.com/Kazunori-Kimura/items/29038632361fba69de5e



## Licensing

This code is licensed under MIT.


## Copyright

2024  [K.Kimura @ Juge.Me](https://github.com/dotnsf) all rights reserved.
