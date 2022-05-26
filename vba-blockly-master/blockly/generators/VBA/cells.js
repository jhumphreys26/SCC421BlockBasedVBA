'use strict';

goog.provide('Blockly.VBA.Cells');

goog.require('Blockly.VBA');

Blockly.VBA['get_cell'] = function (block) {
  var value_prop = Blockly.VBA.valueToCode(block, 'PROP', Blockly.VBA.ORDER_ATOMIC);
  var value_row = Blockly.VBA.valueToCode(block, 'ROW', Blockly.VBA.ORDER_ATOMIC);
  var value_col = Blockly.VBA.valueToCode(block, 'COL', Blockly.VBA.ORDER_ATOMIC);
  var code = '.Cells(' + value_row + ', ' + value_col + ')' + value_prop + '\n';
  return [code, Blockly.VBA.ORDER_ATOMIC];
};

Blockly.VBA['get_cell_worksheet'] = function (block) {
  var value_prop = Blockly.VBA.valueToCode(block, 'PROP', Blockly.VBA.ORDER_ATOMIC);
  var value_row = Blockly.VBA.valueToCode(block, 'ROW', Blockly.VBA.ORDER_ATOMIC);
  var value_col = Blockly.VBA.valueToCode(block, 'COL', Blockly.VBA.ORDER_ATOMIC);
  var value_worksheet = Blockly.VBA.valueToCode(block, 'SHEETNAME', Blockly.VBA.ORDER_ATOMIC);
  var code = 'Worksheets(' + value_worksheet + ').Cells(' + value_row + ',' + value_col + ')' + value_prop + '\n';
  return [code, Blockly.VBA.ORDER_ATOMIC];
};

Blockly.VBA['set_cell_worksheet'] = function (block) {
  var value_prop = Blockly.VBA.valueToCode(block, 'PROP', Blockly.VBA.ORDER_ATOMIC);
  var value_row = Blockly.VBA.valueToCode(block, 'ROW', Blockly.VBA.ORDER_ATOMIC);
  var value_col = Blockly.VBA.valueToCode(block, 'COL', Blockly.VBA.ORDER_ATOMIC);
  var value_worksheet = Blockly.VBA.valueToCode(block, 'WORKSHEET', Blockly.VBA.ORDER_ATOMIC);
  var value_input = Blockly.VBA.valueToCode(block, 'VAL', Blockly.VBA.ORDER_ATOMIC);
  //var code = 'Worksheets("' + value_worksheet + '").Cells(' + value_row + ', ' + value_col + ')' + value_prop + '\n';

  var code = 'ThisWorkbook.Worksheets(' + value_worksheet + ').Cells(' + value_row + ',' + value_col + ')' + value_prop + ' = ' + value_input + '\n';
  return code;
};

Blockly.VBA['set_cell'] = function (block) {
  var value_prop = Blockly.VBA.valueToCode(block, 'PROP', Blockly.VBA.ORDER_ATOMIC);
  var value_row = Blockly.VBA.valueToCode(block, 'ROW', Blockly.VBA.ORDER_ATOMIC);
  var value_col = Blockly.VBA.valueToCode(block, 'COL', Blockly.VBA.ORDER_ATOMIC);
  var value_val = Blockly.VBA.valueToCode(block, 'VAL', Blockly.VBA.ORDER_ATOMIC);
  var code = '.Cells(' + value_row + ', ' + value_col + ')' + value_prop + ' = ' + value_val + '\n';
  return code;
};

Blockly.VBA['value'] = function (block) {
  var code = '.Value';
  return [code, Blockly.VBA.ORDER_ATOMIC];
};

Blockly.VBA['formula'] = function (block) {
  var code = '.Formula';
  return [code, Blockly.VBA.ORDER_ATOMIC];
};

Blockly.VBA['formulaR1C1'] = function (block) {
  var code = '.FormulaR1C1';
  return [code, Blockly.VBA.ORDER_ATOMIC];
};

Blockly.VBA['font'] = function (block) {
  var dropdown_opt = block.getFieldValue('OPT');
  if (dropdown_opt == 'NAME') {
    dropdown_opt = 'Name';
  } else {
    dropdown_opt = 'Size';
  };
  var code = '.Font.' + dropdown_opt;
  return [code, Blockly.VBA.ORDER_ATOMIC];
};

Blockly.VBA['colour'] = function (block) {
  var code = '.Interior.Color'
  return [code, Blockly.VBA.ORDER_ATOMIC];
};

Blockly.VBA['wrap'] = function (block) {
  var code = '.WrapText'
  return [code, Blockly.VBA.ORDER_ATOMIC];
}

Blockly.VBA['fill'] = function (block) {
  var colour_col = block.getFieldValue('COL');
  //convert to RGB
  var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(colour_col)
  var r = parseInt(result[1], 16);
  var g = parseInt(result[2], 16);
  var b = parseInt(result[3], 16);
  var code = 'RGB(' + r + ', ' + g + ', ' + b + ')';
  return [code, Blockly.VBA.ORDER_ATOMIC];
};

Blockly.VBA['border'] = function (block) {
  var dropdown_opt = block.getFieldValue('OPT');
  if (dropdown_opt == "XLDIAGONALDOWN") {
    dropdown_opt = "xlDiagonalDown";
  }
  if (dropdown_opt == "XLDIAGONALUP") {
    dropdown_opt = "xlDiagonalUp";
  }
  if (dropdown_opt == "XLEDGEBOTTOM") {
    dropdown_opt = "xlEdgeBottom";
  }
  if (dropdown_opt == "XLEDGELEFT") {
    dropdown_opt = "xlEdgeLeft";
  }
  if (dropdown_opt == "XLEDGERIGHT") {
    dropdown_opt = "xlEdgeRight";
  }
  if (dropdown_opt == "XLEDGETOP") {
    dropdown_opt = "xlEdgeTop";
  }
  if (dropdown_opt == "XLINSIDEHORIZONTAL") {
    dropdown_opt = "xlInsideHorizontal";
  }
  if (dropdown_opt == "XLINSIDEVERTICAL") {
    dropdown_opt = "xlInsideVertical";
  }

  var code = ".Borders(" + dropdown_opt + ")";
  return [code, Blockly.VBA.ORDER_ATOMIC];
};

Blockly.VBA['linestyle'] = function (block) {
  var dropdown_opt = block.getFieldValue('OPT');
  if (dropdown_opt == "XLCONTINUOUS") {
    dropdown_opt = "xlContinuous";
  }
  if (dropdown_opt == "XLDASH") {
    dropdown_opt = "xldash";
  }
  if (dropdown_opt == "XLDASHDOT") {
    dropdown_opt = "xlDashDot";
  }
  if (dropdown_opt == "XLDASHDOTDOT") {
    dropdown_opt = "xlDashDotDot";
  }
  if (dropdown_opt == "XLDOT") {
    dropdown_opt = "xlDot";
  }
  if (dropdown_opt == "XLDOUBLE") {
    dropdown_opt = "xlDouble";
  }
  if (dropdown_opt == "XLLINESTYLENONE") {
    dropdown_opt = "XlLineStyleNone";
  }
  if (dropdown_opt == "XLSLANTDASHDOT") {
    dropdown_opt = "xlSlantDashDot";
  }

  var code = ".LineStyle = XlLineStyle." + dropdown_opt;
  return [code, Blockly.VBA.ORDER_ATOMIC];
};