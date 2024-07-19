//. app.js
var XLSX = require( 'xlsx' );
var Utils = XLSX.utils;
var eaw = require( 'eastasianwidth' );

var _filepath = null;
var _sheetname = null;
var _row_max_width = 20;
var _border = 1;  //. true
var _value = 1;  //. true
for( var i = 2; i < process.argv.length; i ++ ){
  if( process.argv[i].startsWith( '--' ) ){
    var tmp = process.argv[i].substr( 2 ).split( '=' );
    if( tmp && tmp.length == 2 ){
      switch( tmp[0] ){
      case 'sheet':
        _sheetname = tmp[1];
        break;
      case 'row_max_width':
        _row_max_width = parseInt( tmp[1] );
        break;
      case 'border':
        _border = parseInt( tmp[1] );
        break;
      case 'value':
        _value = parseInt( tmp[1] );
        break;

      default:
      }
    }
  }else if( i == process.argv.length - 1 && !process.argv[i].startsWith( '-' ) ){
    _filepath = process.argv[i];
  }
}

if( _filepath ){
  var book = XLSX.readFile( _filepath );
  
  //. sheets = { Sheet1: {}, Sheet2: {}, .. }
  var sheets = book.Sheets;
  Object.keys( sheets ).forEach( function( sheetname ){
    if( !_sheetname || _sheetname == sheetname ){
      var sheet = sheets[sheetname]
      var cells = [];

      var range = sheet["!ref"];
      var decodeRange = Utils.decode_range( range );  //. { s: { c:0, r:0 }, e: { c:5, r:4 } }

      //. シート内の全セル値を取り出し
      for( var r = decodeRange['s']['r']; r <= decodeRange['e']['r']; r ++ ){
        var row = [];
        for( var c = decodeRange['s']['c']; c <= decodeRange['e']['c']; c ++ ){
          //console.log( {r}, {c} );
          var address = Utils.encode_cell( { r: r, c: c } );
          var cell = sheet[address];
          //console.log( {cell} );
          if( typeof cell !== "undefined" ){
            if( _value ){
              if( typeof cell.v != "undefined" ){
                row.push( cell.v );
              }else{
                row.push( '' );
              }
            }else{
              if( typeof cell.f != "undefined" ){
                if( cell.f ){
                  row.push( cell.f );
                }else{
                  row.push( cell.v );
                }
              }else{
                if( typeof cell.v != "undefined" ){
                  row.push( cell.v );
                }else{
                  row.push( '' );
                }
              }
            }
          }else{
            row.push( '' );
          }
        }
        cells.push( row );
      }

      if( cells && cells.length ){
        //console.log( cells );
        //. シート内の列毎の最大長さを定義
        var row_widths = [];
        
        //. 初期化
        for( var c = 0; c < cells[0].length; c ++ ){
          row_widths.push( 0 );
        }

        for( var r = 0; r < cells.length; r ++ ){
          for( var c = 0; c < cells[r].length; c ++ ){
            var cell = cells[r][c]
            var str_cell = '' + cell;
            //console.log( {r}, {c}, {str_cell} );
            if( row_widths[c] < eaw.length( str_cell ) ){
              if( _row_max_width < eaw.length( str_cell ) ){
                row_widths[c] = _row_max_width;
              }else{
                row_widths[c] = eaw.length( str_cell );
              }
            }
          }
        }
        //console.log( row_widths );

        //. Display
        displaySheet( sheetname, cells, row_widths, _border, _row_max_width );
      }
    }
  });
}

function displaySheet( s, cl, rw, b, rmw ){
  //console.log( {s} );
  //console.log( {cl} );
  //console.log( {rw} );
  //console.log( {b} );
  //console.log( {rmw} );
  console.log( s + ' :' );

  if( _border ){
    displayBorderLine( rw );
  }

  for( var r = 0; r < cl.length; r ++ ){
    //. c[r] 行を出力するために何行必要か？
    //. c[r] 行を出力するために必要な行数(mx) = 各 c[r][c] セルを出力するために必要な行数の最大値
    var mx = 1;
    for( var c = 0; c < cl[r].length; c ++ ){
      var cell = cl[r][c];
      var str_cell = '' + cell;
      var cell_length = eaw.length( str_cell );
      var h = Math.ceil( cell_length / rw[c] );
      if( h > mx ){ mx = h; }
    }

    //. c[r] 行を mx 行で表示する
    var line_values = [];
    for( var c = 0; c < cl[r].length; c ++ ){
      var cell = cl[r][c];
      var str_cell = '' + cell;
      var cell_values = [];  //. これの要素数が mx になる

      var idx = 0;
      for( var i = 0; i < mx; i ++ ){
        var s = eawSubstr( str_cell, idx, rw[c] );
        idx += eaw.length( s );
        while( eaw.length( s ) < rw[c] ){
          s += ' ';
        }
        cell_values.push( s );
      }

      line_values.push( cell_values );
    }

    //. cells[r] 行を mx 行使って出力する
    for( var i = 0; i < mx; i ++ ){
      var line = '';
      if( _border ){
        line += '|';
      }
      for( var j = 0; j < line_values.length; j ++ ){
        line += ' ';
        line += line_values[j][i];
        line += ' ';
        if( _border ){
          line += '|';
        }
      }

      console.log( line );
    }

    if( _border ){
      displayBorderLine( rw );
    }
  }

  console.log( '' );
}

function displayBorderLine( rw ){
  var line = '+';
  for( var c = 0; c < rw.length; c ++ ){
    for( var i = 0; i < rw[c] + 2; i ++ ){  //. 左右に空白１つぶん余計に出力する
      line += '-';
    }

    line += '+';
  }

  console.log( line );
}

function eawSubstr( str, begin, num, fill = false ){
  var b = false;
  var t = '';
  var r = '';
  var len = str.length;
  for( var i = 0; i < len && !b; i ++ ){
    var c = str.charAt( i );
    if( eaw.length( t ) >= begin ){
      if( eaw.length( r + c ) <= num ){
        r += c;
      }else{
        b = true;
      }
    }else{
      t += c;
    }
  }

  if( fill ){
    while( eaw.length( r ) < num ){
      r += fill;
    }
  }

  return r;
}
