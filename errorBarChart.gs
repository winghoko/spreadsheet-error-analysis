"use strict";
/** @OnlyCurrentDoc */

/**
 * @license MIT
 * 
 * MIT License
 *
 * Copyright (c) 2024 winghoko
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

/**
 * Add new menu item for error bar chart
 */
function onOpen() {
  // Get the Ui object. 
  var ui = SpreadsheetApp.getUi();
  
  try { // branch depending on whether WLINEST.gs exists
    exist_WLINEST;
    // Create a custom menu. 
    ui.createMenu('ErrorAnalysis')
      .addItem('Make Chart...', "createErrorChart")
      .addItem('Add to Chart...', "appendErrorChart")
      .addSeparator()
      .addItem('Help on "Make Chart"', "helpMakeChart")
      .addItem('Help on "Add to Chart"', "helpAppendChart")
      .addItem('Help on "WLINEST()"', "helpWLINEST")
      .addToUi();
  } catch(e){
      ui.createMenu('ErrorAnalysis')
        .addItem('Make Chart...', "createErrorChart")
        .addItem('Add to Chart...', "appendErrorChart")
        .addSeparator()
        .addItem('Help on "Make Chart"', "helpMakeChart")
        .addItem('Help on "Add to Chart"', "helpAppendChart")
        .addToUi();
  }
}

function createErrorChart() {

  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var response;

  var title = "Scatter plot with error bars", y_label = "known_data_y", x_label = "known_data_x";
  var num_rows = 0, ins_Col, ins_Row;
  var y_range = undefined, x_range = undefined, u_values = undefined;
  var activeRange = sheet.getActiveRange();

  if (activeRange !== null && activeRange.getNumColumns() === 3){ // assume standard data layout

    if ( isNaN(Number(activeRange.getValue())) ){ // assume first row is header
      x_label = activeRange.getValue();
      y_label = activeRange.offset(0, 1).getValue();
      activeRange = activeRange.offset(1, 0, activeRange.getNumRows() - 1);
    }

    // extract data references, values, and properties
    num_rows = activeRange.getNumRows();
    x_range = activeRange.offset(0, 0, num_rows, 1);
    y_range = activeRange.offset(0, 1, num_rows, 1);
    u_values = activeRange.offset(0, 2, num_rows, 1).getValues().map(v => Number(v));

    // location at which the chart will be inserted
    ins_Row = activeRange.getRow() + num_rows + 1;
    ins_Col = activeRange.getColumn();

  } else { // explicitly asked for data locations

    // fall back on location to insert chart
    if (activeRange === null){
      activeRange = sheet.getCurrentCell();
    }

    // location at which the chart will be inserted
    ins_Col = activeRange.getColumn();
    ins_Row = activeRange.getRow();

    // extract x-data
    var x_range = undefined;
    response = ui.prompt(
      'known_data_x', 
      'Enter the cell range for independent (x) variable', 
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() == ui.Button.OK){
      x_range = response.getResponseText();
      x_range = sheet.getRange(x_range);
    }

    // extract y-data (and number of rows)
    var y_range = undefined;
    response = ui.prompt(
      'known_data_y', 
      'Enter the cell range for dependent (y) variable', 
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() == ui.Button.OK){
      y_range = response.getResponseText();
      y_range = sheet.getRange(y_range);
    }
    num_rows = y_range.getNumRows();

    // extract uncertainty data
    var u_values = undefined;
    response = ui.prompt(
      'uncertainty_y', 
      'Enter the cell range for uncertainty in dependent variable', 
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() == ui.Button.OK){
      u_values = response.getResponseText();
      u_values = sheet.getRange(u_values).getValues().map(v => Number(v));
    } 

  }

  // ask for color for error bars
  var color;
  response = ui.prompt(
    'color', 
    'Enter the color for error bars (default = black)', 
    ui.ButtonSet.OK
  );
  color = response.getResponseText();
  color = (color==="") ? '#000000' : color

  // NOTE: "String(v) !== ''" seems to be the best way to filter
  // out "empty cell" without filtering out 0's

  // create auxillary sheet
  var i = 1; 
  while (spreadsheet.getSheetByName("error_bars_" + i) !== null ){
    i++;
  }
  var aux_name = "error_bars_" + i
  var aux_sheet = spreadsheet.insertSheet().setName(aux_name);
  var col_idxes = [];
  x_range.copyTo(aux_sheet.getRange(1,1), {contentsOnly: true})
  y_range.getValues().forEach( (v, i) => {
    if (String(v) !== ''){
      aux_sheet.getRange(1+i, 2 + col_idxes.length).setValue(v);
      col_idxes.push(i);
    }
  })
  sheet.activate();

  // create chart; provide info for x-axis
  var chart = sheet.newChart().asScatterChart().addRange(x_range);

  // add a data series with a single error bar for each y value
  col_idxes.forEach( (v, i) => {
    chart = chart.addRange(aux_sheet.getRange(1, i + 2, num_rows + 1))
      .setOption('series.' + i + '.errorBars.magnitude', u_values[v])
      .setOption('series.' + i + '.errorBars.errorType', 'constant')
      .setOption('series.' + i + '.color', color)
      .setOption('series.' + i + '.pointSize', 0)
  })

  // add the last series: the actual y-data
  chart = chart.addRange(y_range)
    .setOption('series.' + col_idxes.length + '.color', color)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setOption('legend.position', 'none')
    .setYAxisTitle(y_label)
    .setXAxisTitle(x_label)
    .setTitle(title)
    .setPosition(ins_Row, ins_Col, 0, 0)
    .build()

  sheet.insertChart(chart);  
}

function appendErrorChart() {

  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var response;

  var num_rows = 0;
  var y_range = undefined, u_values = undefined;
  var activeRange = sheet.getActiveRange();

  if (activeRange !== null && activeRange.getNumColumns() === 2){ // assume standard data layout

    if ( isNaN(Number(activeRange.getValue())) ){ // assume first row is header
      activeRange = activeRange.offset(1, 0, activeRange.getNumRows() - 1);
    }

    // extract data reference, values, and properties
    num_rows = activeRange.getNumRows();
    y_range = activeRange.offset(0, 0, num_rows, 1);
    u_values = activeRange.offset(0, 1, num_rows, 1).getValues().map(v => Number(v));

  } else { // explicitly asked for data locations

    // extract y-data
    var y_range = undefined;
    response = ui.prompt(
      'known_data_y', 
      'Enter the cell range for dependent (y) variable', 
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() == ui.Button.OK){
      y_range = response.getResponseText();
      y_range = sheet.getRange(y_range);
    }
    num_rows = y_range.getNumRows();

    // extract uncertainty data
    var u_values = undefined;
    response = ui.prompt(
      'uncertainty_y', 
      'Enter the cell range for uncertainty in dependent variable', 
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() == ui.Button.OK){
      u_values = response.getResponseText();
      u_values = sheet.getRange(u_values).getValues().map(v => Number(v));
    } 

  }

  // hook to the last created chart
  var chart = sheet.getCharts();
  var chart = chart[chart.length - 1]
  var last = chart.getRanges().length - 1;
  var x_range = chart.getRanges()[0]

  // ask for color for data points
  var color;
  response = ui.prompt(
    'color', 
    'Enter the color for data points (default = black)', 
    ui.ButtonSet.OK
  );
  color = response.getResponseText();
  color = (color==="") ? '#000000' : color

  // NOTE: "String(v) !== ''" seems to be the best way to filter
  // out "empty cell" without filtering out 0's
  
  // create auxillary sheet
  var i = 1; 
  while (spreadsheet.getSheetByName("error_bars_" + i) !== null ){
    i++;
  }
  var aux_name = "error_bars_" + i
  var aux_sheet = spreadsheet.insertSheet().setName(aux_name);
  var col_idxes = [];
  x_range.copyTo(aux_sheet.getRange(1,1), {contentsOnly: true})
  y_range.getValues().forEach( (v, i) => {
    if (String(v) !== ''){
      aux_sheet.getRange(1+i, 2 + col_idxes.length).setValue(v);
      col_idxes.push(i);
    }
  })
  sheet.activate();

  // add a data series with a single error bar for each y value
  var idx = 0;
  chart = chart.modify();
  col_idxes.forEach( (v, i) => {
    idx = last + i;
    chart = chart.addRange(aux_sheet.getRange(1, i + 2, num_rows + 1))
      .setOption('series.' + idx + '.errorBars.magnitude', u_values[v])
      .setOption('series.' + idx + '.errorBars.errorType', 'constant')
      .setOption('series.' + idx + '.color', color)
      .setOption('series.' + idx + '.pointSize', 0)
  })
  // add the last series: the actual y-data
  idx = last + num_rows;
  chart = chart.addRange(y_range)
    .setOption('series.' + idx + '.color', color)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setOption('legend.position', 'none')
    .build()

  sheet.updateChart(chart);
}

var make_chart_msg = `Create a chart with error bars whose magnitude vary with data.

Expects 3 input columns of the same height: the values of the independent ("X") variable, the dependent ("Y") variable, and the uncertainty ("u(Y)") in the dependent variable.

If a 3-column range of cells is selected, the range is assumed to be the input data, in the order of X, Y, u(Y). In addition, the first row will be interpreted as header if it cannot be parsed as numbers.

Otherwise, the user must input each data column separately using the standard "A1"
notation (e.g., A2:A10), with header omitted.

An auxillary sheet is created for the purpose of the plot. Deleting this auxillary sheet will mess up the plot.

Because some data values are hard-coded in the chart, to modify chart data you MUST rerun the "Make Chart..." option again after the data is updated.`

function helpMakeChart(){

  SpreadsheetApp.getUi().alert(
    'Help on make error-bar chart',
    make_chart_msg,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

var append_chart_msg = `Add an extra data series (with data-dependent error bars) to existing chart.

Only the LAST chart in the CURRENT sheet is modified. And the new data series must take the same values in the independent ("X") variable as one already plotted.

Expects 2 input columns: the values of dependent ("Y") variable and its uncertainty ("u(Y)").

If a 2-column range of cells is selected, the range is assumed to be the input data, in the order of Y, u(Y). In addition, the first row will be interpreted as header if it cannot be parsed as numbers.

Otherwise, the user must input each data column separately using the standard "A1"notation (e.g., A2:A10), with header omitted.

An (additional) auxillary sheet is created for the purpose of the plot. Deleting this auxillary sheet will mess up the plot.`

function helpAppendChart(){

  SpreadsheetApp.getUi().alert(
    'Help on adding data to error-bar chart',
    append_chart_msg,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}