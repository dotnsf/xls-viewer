//. app.js
var XLSX = require( 'xlsx' );
var Utils = XLSX.utils;
var eaw = require( 'eastasianwidth' );

var _filepath = null;
var _sheetnames = null;
var _row_max_width = 20;
var _border = 1;   //. true
var _formula = 0;  //. false
var _a1 = 0;       //. false #2
var _label = 0;    //. false #1
for( var i = 2; i < process.argv.length; i ++ ){
  if( process.argv[i].startsWith( '--' ) ){
    var tmp = process.argv[i].substr( 2 ).split( '=' );
    if( tmp && tmp.length == 2 ){
      switch( tmp[0] ){
      case 'sheets':
        _sheetnames = tmp[1].split( ',' );
        break;
      case 'row_max_width':
        _row_max_width = parseInt( tmp[1] );
        break;
      case 'border':
        _border = parseInt( tmp[1] );
        break;
      case 'formula':
        _formula = parseInt( tmp[1] );
        break;
      case 'a1':
        _a1 = parseInt( tmp[1] );
        break;
      case 'label':
        _label = parseInt( tmp[1] );
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
    if( !_sheetnames || _sheetnames.length == 0 || _sheetnames.indexOf( sheetname ) > -1 ){
      var sheet = sheets[sheetname]
      var cells = [];

      var range = sheet["!ref"];
      var decodeRange = Utils.decode_range( range );  //. { s: { c:0, r:0 }, e: { c:5, r:4 } }
      //. #2
      if( _a1 ){
        decodeRange['s']['c'] = 0;
        decodeRange['s']['r'] = 0;
      }

      //. #1
      if( _label ){
        var row = [ '' ];
        var codeA = 65;  // 'A'
        var codeZ = 90;  // 'Z'
        for( var c = decodeRange['s']['c']; c <= decodeRange['e']['c']; c ++ ){
          var code = c;  //. 26 以上の可能性あり
          var column_label = '';
          while( code >= 0 ){
            if( code > 25 ){
              var t1 = code % 26;
              var t2 = Math.floor( code / 26 );

              column_label = String.fromCharCode( codeA + t1 ) + column_label;
              code = t2 - 1;
            }else{
              column_label = String.fromCharCode( codeA + code ) + column_label;
              code = -1;
            }
          }
          row.push( column_label );
        }
        cells.push( row );
      }

      //. シート内の全セル値を取り出し
      for( var r = decodeRange['s']['r']; r <= decodeRange['e']['r']; r ++ ){
        var row = [];

        //. #1
        if( _label ){
          row.push( r + 1 );
        }

        for( var c = decodeRange['s']['c']; c <= decodeRange['e']['c']; c ++ ){
          var address = Utils.encode_cell( { r: r, c: c } );
          var cell = sheet[address];
          if( typeof cell !== "undefined" ){
            if( !_formula ){
              if( typeof cell.v != "undefined" ){
                row.push( cell.v );
              }else{
                row.push( '' );
              }
            }else{
              if( typeof cell.f != "undefined" ){
                row.push( cell.f );
              }else{
                row.push( '' );
              }
            }
          }else{
            row.push( '' );
          }
        }
        cells.push( row );
      }

      if( cells && cells.length ){
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
            if( row_widths[c] < eaw.length( str_cell ) ){
              if( _row_max_width < eaw.length( str_cell ) ){
                row_widths[c] = _row_max_width;
              }else{
                row_widths[c] = eaw.length( str_cell );
              }
            }
          }
        }

        //. Display
        displaySheet( sheetname, cells, row_widths, _border, _row_max_width, _label );
      }
    }
  });
}

function displaySheet( s, cl, rw, b, rmw, l ){
  console.log( s + ' :' );

  if( b ){
    displayBorderLine( rw, l );
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

    if( b ){
      if( r == 0 ){
        displayBorderLine( rw, l );
      }else{
        displayBorderLine( rw );
      }
    }
  }

  console.log( '' );
}

function displayBorderLine( rw, l ){
  var line = '+';
  for( var c = 0; c < rw.length; c ++ ){
    for( var i = 0; i < rw[c] + 2; i ++ ){  //. 左右に空白１つぶん余計に出力する
      if( l || ( _label && c == 0 ) ){
        line += '=';  //. #1
      }else{
        line += '-';
      }
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
