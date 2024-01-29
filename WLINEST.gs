"use strict";

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

var exist_WLINEST = true;

/**
 * Weighted least-square fit for best linear trend.
 * 
 * If verbose is FALSE, returns the row [fitted slope, fitted intercept].
 * 
 * If verbose is true, returns the 3-by-2 table [ [fitted slope, fitted intercept], 
 * [slope error, intercept error], [unscaled chi-square, degree of freedom] ]
 *
 * @param {Array<number>} known_data_y - The values or range of the dependent (y) variables already known.
 * 
 * @param {Array<number>} known_data_x - The values or range of the independent (x) variable 
 * corresponding to known_data_y.
 * 
 * @param {Array<number>} uncertainty_y - The values or range of uncertainty (standard deviation)
 * corresponding to known_data_y.
 * 
 * @param {boolean} calculate_b - [Optional, default=TRUE] if TRUE, the y-intercept (b) is calculated; 
 * otherwise it is set at 0.
 * 
 * @param {boolean} absolute - [Optional, default=FALSE] if TRUE, uncertainty are treated as absolute; 
 * otherwise, they are treated as relative ratios. (NOTE: absolute affects only the error in the fitted
 * parameters).
 * 
 * @param {boolean} verbose - [Optional, default=FALSE] if TRUE, return additional regression statistics;
 *  otherwise return only fitted values.
 * 
 * @returns if verbose is FALSE, the row [fitted slope, fitted intercept]; if verbose is true, 
 * the 3-by-2 table [ [fitted slope, fitted intercept], [slope error, intercept error],
 * [unscaled weighed chi-square, degree of freedom] ]
 * 
 * @customfunction
 */
function WLINEST(known_data_y, known_data_x, uncertainty_y, calculate_b=true, absolute=false, verbose=false) {

  // coerce input into number type
  known_data_y = known_data_y.map(v => Number(v))
  known_data_x = known_data_x.map(v => Number(v))
  uncertainty_y = uncertainty_y.map(v => Number(v))

  // various weighed sums, computed in a single loop
  var sum_w = 0, sum_wx = 0, sum_wxx = 0, sum_wy = 0, sum_wxy = 0;
  uncertainty_y.forEach( (val, idx) => {
      var w = 1/val/val;
      var x = known_data_x[idx];
      var y = known_data_y[idx];
      sum_w += w;
      sum_wx += w * x;
      sum_wxx += w * x * x;
      sum_wy += w * y;
      sum_wxy += w * x * y;
  })

  // fitted slope and intercept
  var inv_det = 1, slope = 0, intercept = 0;
  if (calculate_b){
    inv_det = 1/(sum_w * sum_wxx - sum_wx * sum_wx);
    slope = inv_det * (sum_w * sum_wxy - sum_wx * sum_wy);
    intercept = inv_det * (sum_wxx * sum_wy - sum_wx * sum_wxy);
  } else {
    inv_det = 1/sum_wxx;
    slope = inv_det * sum_wxy;
  }

  // early return if only fitted parameters are asked for
  if (!verbose){
    return [[slope, intercept]]
  }

  // compute weighed chi-square assuming y uncertainties absolute
  var chisq = 0.0;
  known_data_y.forEach( (val, idx) => {
    var x = known_data_x[idx];
    var w = uncertainty_y[idx];
    w = 1/w/w;
    chisq += w * Math.pow(val - intercept - slope * x, 2);
  });
  // compute degree of freedom
  var df = calculate_b ? known_data_y.length - 2 : known_data_y.length - 1

  // compute uncertainties in fitted parameters
  var u_slope = 1, u_intercept = Math.nan;
  if (calculate_b){
    u_intercept = Math.sqrt(inv_det * sum_wxx);
    u_slope = Math.sqrt(inv_det * sum_w);
  } else {
    u_slope = Math.sqrt(inv_det);
  }

  // adjust uncertainties in fitted parameter when uncertain_y are relative
  if (!absolute){
    u_slope *= Math.sqrt(chisq/df);
    u_intercept *= Math.sqrt(chisq/df);
  }

  return [
    [slope, intercept],
    [u_slope, u_intercept],
    [chisq, df]
  ]

}

var help_WLINEST_msg = `The WLINEST() function performs weighted least square fit for best linear trend, and its interface is designed to mimic the built-in LINEST() function.

To use, simply enter "=WLINEST(...)" in a cell, with ... the appropriate arguments. To the minimum, the values of dependent ("Y") variable, independent ("X") variable, and uncertainty ("u(Y)") in dependent varaible are needed, and these can be supplied as spreadsheet ranges (e.g., "A2:A20").

At present the function can only be used with ONE independent variable. Hence, Y, X, and u(Y) are all expected to be a single column, and should have the same height.

Note also that the WLINEST() function returns results by writing in MULTIPLE cells (1-by-2 or 3-by-2 depending on settings), and that it has the usual user-defined function restriction that it can only operates on deterministic data.

For more information, type "=WLINEST(" into the formula bar and expand its tooltip by hovering/clicking the "v" or "?" symbols that shows up.`

function helpWLINEST(){

  SpreadsheetApp.getUi().alert(
    'Help on the WLINEST() function',
    help_WLINEST_msg,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}