// =================================================================
// Configuration
// =================================================================
var INDICATORS_SPREADSHEET_ID = '1uHxxR63wdA7CyBaRpji2dy8A4AFQ3zK7AbkuFZm_kgI';
var RESULTS_SPREADSHEET_ID = '1TicIR6zJRNDWq4gOFN1vKwFKh7DVvX4RDMS-hXuH9io';
var USERS_SPREADSHEET_ID = '1KoEfTGx2HH062Wa_vF1byE3GlJ-ExtrB94mTeZilVSo';

var SHEET_NAMES = {
  INDICATORS: 'Indicators',
  RESULTS: 'Results', 
  USERS: 'Users'
};

// =================================================================
// Helper Functions
// =================================================================
function getSheetById(spreadsheetId, sheetName) {
  try {
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw new Error('Sheet not found: ' + sheetName);
    return sheet;
  } catch (e) {
    throw new Error('Failed to access spreadsheet: ' + e.message);
  }
}

function getCachedData(key, fetchFunction, expirationInMinutes) {
  expirationInMinutes = expirationInMinutes || 5;
  var cache = CacheService.getScriptCache();
  var cachedData = cache.get(key);
  if (cachedData != null) return JSON.parse(cachedData);
  var freshData = fetchFunction();
  cache.put(key, JSON.stringify(freshData), expirationInMinutes * 60);
  return freshData;
}

function _mapRowToIndicator(row) {
  return { 
    department: row[0] ? String(row[0]).trim() : '',
    name: row[1] ? String(row[1]).trim() : '',
    description: row[2] ? String(row[2]).trim() : '',
    target: row[11] ? String(row[11]).trim() : '',
    benchmark: row[12] ? String(row[12]).trim() : ''
  };
}

function _mapRowToUser(row) {
  return { 
    email: row[0] ? String(row[0]).trim() : '',
    department: row[1] ? String(row[1]).trim() : '',
    role: row[2] ? String(row[2]).trim() : ''
  };
}

function _mapRowToResult(row) {
  var monthValue = row[2] ? String(row[2]).trim() : '';
  
  return { 
    department: row[0] ? String(row[0]).trim() : '',
    indicator: row[1] ? String(row[1]).trim() : '',
    month: monthValue,
    monthFormatted: formatMonthForDisplay(monthValue),
    value: row[3] ? String(row[3]).trim() : '',
    numerator: row[4] ? String(row[4]).trim() : '',
    denominator: row[5] ? String(row[5]).trim() : '',
    result: row[6] ? String(row[6]).trim() : '',
    note: row[7] ? String(row[7]).trim() : ''
  };
}

// =================================================================
// Date Formatting Functions
// =================================================================
function formatMonthForDisplay(monthString) {
  if (!monthString) return '';
  
  try {
    // إذا كان التاريخ يظهر ككائن Date
    if (typeof monthString === 'object' || monthString.toString().includes('GMT')) {
      var date = new Date(monthString);
      var month = date.getMonth() + 1;
      var year = date.getFullYear();
      return formatShortMonth(month, year);
    }
    
    // إذا كان بصيغة YYYY-MM
    if (monthString.includes('-')) {
      var parts = monthString.split('-');
      if (parts.length === 2) {
        var year = parts[0];
        var month = parts[1];
        return formatShortMonth(parseInt(month), parseInt(year));
      }
    }
    
    return monthString;
  } catch (error) {
    return monthString;
  }
}

function formatShortMonth(month, year) {
  var monthNames = {
    1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
    7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
  };
  return monthNames[month] + ' ' + year;
}

// =================================================================
// Main Display
// =================================================================
function doGet(e) {
  var page = e && e.parameter ? e.parameter.page : 'main';
  var htmlFile = 'Main';
  
  switch(page) {
    case 'dashboard': htmlFile = 'Dashboard'; break;
    case 'admin': htmlFile = 'Admin'; break;
    case 'users': htmlFile = 'Users'; break;
    default: htmlFile = 'Main';
  }

  try {
    var template = HtmlService.createTemplateFromFile(htmlFile);
    template.appUrl = ScriptApp.getService().getUrl();
    return template.evaluate()
        .setTitle('KPI Platform')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    return HtmlService.createHtmlOutput('<h1>Error Loading Page</h1><p>' + error.message + '</p>');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =================================================================
// Auth & Users
// =================================================================
function getCurrentUserRole() {
  try {
    var email = Session.getActiveUser().getEmail();
    var usersSheet = getSheetById(USERS_SPREADSHEET_ID, SHEET_NAMES.USERS);
    var data = usersSheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) {
        return { 
          role: data[i][2] ? String(data[i][2]).trim() : 'guest',
          email: email,
          department: data[i][1] ? String(data[i][1]).trim() : ''
        };
      }
    }
    return { role: 'guest', email: email, department: '' };
  } catch (error) {
    return { role: 'guest', email: Session.getActiveUser().getEmail(), department: '' };
  }
}

function doLogout() {
  return { success: true, message: 'Logged out successfully' };
}

// =================================================================
// Indicators Management
// =================================================================
function getAllIndicators() {
  return getCachedData('ALL_INDICATORS_CACHE', _fetchAllIndicatorsFromSheet, 5);
}

function _fetchAllIndicatorsFromSheet() {
  try {
    var sheet = getSheetById(INDICATORS_SPREADSHEET_ID, SHEET_NAMES.INDICATORS);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    var indicators = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0] && row[1]) {
        indicators.push(_mapRowToIndicator(row));
      }
    }
    return indicators;
  } catch (error) {
    console.error('Error fetching indicators:', error);
    return [];
  }
}

function getIndicatorByName(name) {
  try {
    var searchName = String(name).trim();
    var allIndicators = getAllIndicators();
    
    for (var i = 0; i < allIndicators.length; i++) {
      if (String(allIndicators[i].name).trim() === searchName) {
        return allIndicators[i];
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting indicator by name:', error);
    return null;
  }
}

function getIndicatorsByDepartment(department) {
  try {
    var searchDepartment = String(department).trim();
    var allIndicators = getAllIndicators();
    return allIndicators.filter(function(ind) {
      return String(ind.department).trim() === searchDepartment;
    });
  } catch (error) {
    console.error('Error getting indicators by department:', error);
    return [];
  }
}

function getDepartments() {
  try {
    var allIndicators = getAllIndicators();
    var departments = [];
    var seen = {};
    
    allIndicators.forEach(function(ind) {
      if (ind.department && !seen[ind.department]) {
        seen[ind.department] = true;
        departments.push(ind.department);
      }
    });
    return departments.sort();
  } catch (error) {
    console.error('Error getting departments:', error);
    return [];
  }
}

// =================================================================
// Results Management - FIXED VERSION
// =================================================================
function getResults(department, indicator) {
  try {
    var searchDepartment = String(department).trim();
    var searchIndicator = String(indicator).trim();
    
    var sheet = getSheetById(RESULTS_SPREADSHEET_ID, SHEET_NAMES.RESULTS);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row.length >= 8 && 
          String(row[0]).trim() === searchDepartment && 
          String(row[1]).trim() === searchIndicator) {
        
        var result = _mapRowToResult(row);
        result.rowIndex = i + 1;
        results.push(result);
      }
    }
    return results;
  } catch (error) {
    console.error('Error getting results:', error);
    return [];
  }
}

function getResultsWithDateRange(department, indicator, startDate, endDate) {
  try {
    var allResults = getResults(department, indicator);
    
    if (!startDate || !endDate) {
      return allResults.sort(function(a, b) {
        var dateA = new Date(a.month + '-01');
        var dateB = new Date(b.month + '-01');
        return dateA - dateB;
      });
    }
    
    var filteredResults = allResults.filter(function(result) {
      try {
        var resultDate = new Date(result.month + '-01');
        var start = new Date(startDate + '-01');
        var end = new Date(endDate + '-01');
        return resultDate >= start && resultDate <= end;
      } catch (e) {
        return true;
      }
    });
    
    return filteredResults.sort(function(a, b) {
      var dateA = new Date(a.month + '-01');
      var dateB = new Date(b.month + '-01');
      return dateA - dateB;
    });
  } catch (error) {
    console.error('Error getting results with date range:', error);
    return getResults(department, indicator);
  }
}

// =================================================================
// Enhanced Duplicate Prevention - FIXED
// =================================================================
function checkDuplicateResult(department, indicator, month) {
  try {
    var sheet = getSheetById(RESULTS_SPREADSHEET_ID, SHEET_NAMES.RESULTS);
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row.length >= 3) {
        var rowDepartment = String(row[0]).trim();
        var rowIndicator = String(row[1]).trim();
        var rowMonth = String(row[2]).trim();
        
        if (rowDepartment === department && 
            rowIndicator === indicator && 
            rowMonth === month) {
          return true;
        }
      }
    }
    return false;
  } catch (error) {
    console.error('Error checking duplicate:', error);
    return false;
  }
}

function calculateAutomaticResult(numerator, denominator, value) {
  if (value && !isNaN(parseFloat(value))) {
    return String(parseFloat(value));
  }
  
  if (numerator && denominator && !isNaN(parseFloat(numerator)) && !isNaN(parseFloat(denominator))) {
    var num = parseFloat(numerator);
    var den = parseFloat(denominator);
    if (den !== 0) {
      var result = (num / den) * 100;
      return String(result.toFixed(2));
    }
  }
  
  return '';
}

function saveResult(data) {
  console.log('Saving result:', data);
  
  if (!data.department || !data.indicator || !data.month) {
    throw new Error('Department, indicator and month are required');
  }
  
  var isDuplicate = checkDuplicateResult(data.department, data.indicator, data.month);
  if (isDuplicate) {
    var existingMonth = formatMonthForDisplay(data.month);
    throw new Error('❌ Result already exists for ' + existingMonth + '. Please use the edit function to update the existing result.');
  }
  
  var sheet = getSheetById(RESULTS_SPREADSHEET_ID, SHEET_NAMES.RESULTS);
  
  var finalResult = data.result;
  if (!finalResult) {
    finalResult = calculateAutomaticResult(data.numerator, data.denominator, data.value);
  }
  
  var resultData = [
    data.department,
    data.indicator,
    data.month,
    data.value || '',
    data.numerator || '',
    data.denominator || '',
    finalResult || '',
    data.note || ''
  ];
  
  sheet.appendRow(resultData);
  
  CacheService.getScriptCache().remove('ALL_INDICATORS_CACHE');
  
  return { 
    success: true, 
    message: '✅ Result saved successfully for ' + formatMonthForDisplay(data.month)
  };
}

function updateResult(rowIndex, data) {
  console.log('Updating result:', data);
  
  try {
    var sheet = getSheetById(RESULTS_SPREADSHEET_ID, SHEET_NAMES.RESULTS);
    
    var finalResult = data.result;
    if (!finalResult) {
      finalResult = calculateAutomaticResult(data.numerator, data.denominator, data.value);
    }
    
    var resultData = [
      data.department,
      data.indicator,
      data.month,
      data.value || '',
      data.numerator || '',
      data.denominator || '',
      finalResult || '',
      data.note || ''
    ];
    
    sheet.getRange(rowIndex, 1, 1, resultData.length).setValues([resultData]);
    
    CacheService.getScriptCache().remove('ALL_INDICATORS_CACHE');
    
    return { 
      success: true, 
      message: '✅ Result updated successfully for ' + formatMonthForDisplay(data.month)
    };
  } catch (error) {
    throw new Error('Failed to update result: ' + error.message);
  }
}

function deleteResult(rowIndex) {
  console.log('Deleting result at row:', rowIndex);
  
  try {
    var sheet = getSheetById(RESULTS_SPREADSHEET_ID, SHEET_NAMES.RESULTS);
    sheet.deleteRow(rowIndex);
    
    CacheService.getScriptCache().remove('ALL_INDICATORS_CACHE');
    
    return { success: true, message: '✅ Result deleted successfully' };
  } catch (error) {
    throw new Error('Failed to delete result: ' + error.message);
  }
}

function getIndicatorWithTargets(indicatorName) {
  try {
    var indicator = getIndicatorByName(indicatorName);
    if (!indicator) throw new Error('Indicator not found');
    
    return {
      indicator: indicator,
      targets: {
        target: indicator.target ? parseFloat(indicator.target) : null,
        benchmark: indicator.benchmark ? parseFloat(indicator.benchmark) : null
      }
    };
  } catch (error) {
    throw error;
  }
}

// =================================================================
// Enhanced PDF Report with Chart
// =================================================================
function sendIndicatorReportEmail(indicatorName) {
  try {
    var indicator = getIndicatorByName(indicatorName);
    if (!indicator) throw new Error('Indicator not found');
    
    var results = getResults(indicator.department, indicator.name);
    
    var sortedResults = results.sort(function(a, b) {
      var dateA = new Date(a.month + '-01');
      var dateB = new Date(b.month + '-01');
      return dateA - dateB;
    });
    
    var htmlContent = generateReportHtmlWithChart(indicator, sortedResults);
    var blob = Utilities.newBlob(htmlContent, 'text/html', 'report.html')
      .getAs('application/pdf')
      .setName('KPI_Report_' + indicatorName + '_' + new Date().toISOString().split('T')[0] + '.pdf');
    
    var file = DriveApp.createFile(blob);
    return 'https://drive.google.com/uc?export=download&id=' + file.getId();
  } catch (error) {
    throw new Error('Report failed: ' + error.message);
  }
}

function generateReportHtmlWithChart(indicator, results) {
  var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  
  var html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <title>KPI Report - ${indicator.name}</title>
    <style>
      body { 
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
        margin: 0;
        padding: 20px;
        line-height: 1.6;
        color: #333;
        background: #f8f9fa;
      }
      .container {
        max-width: 1000px;
        margin: 0 auto;
        background: white;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      }
      .header { 
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 30px;
        border-radius: 10px;
        margin-bottom: 30px;
        text-align: center;
      }
      .header h1 {
        margin: 0;
        font-size: 28px;
        font-weight: 600;
      }
      .header h2 {
        margin: 10px 0 0 0;
        font-size: 20px;
        font-weight: 400;
        opacity: 0.9;
      }
      .info-section {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 20px;
        margin-bottom: 30px;
      }
      .indicator-info, .targets { 
        background: #f8f9fa; 
        padding: 20px; 
        border-radius: 8px;
        border-left: 4px solid #667eea;
      }
      .targets {
        border-left-color: #28a745;
      }
      .section-title {
        font-size: 20px;
        font-weight: 600;
        color: #2c3e50;
        margin: 25px 0 15px 0;
        padding-bottom: 8px;
        border-bottom: 2px solid #667eea;
      }
      .stats {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
        gap: 15px;
        margin: 20px 0;
      }
      .stat-box {
        background: white;
        padding: 20px;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-top: 4px solid #667eea;
      }
      .stat-value {
        font-size: 24px;
        font-weight: 700;
        color: #667eea;
        margin-bottom: 5px;
      }
      .stat-label {
        font-size: 12px;
        color: #666;
        text-transform: uppercase;
        letter-spacing: 0.5px;
      }
      .chart-container {
        background: white;
        padding: 25px;
        border-radius: 8px;
        margin: 25px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
      }
      .chart-title {
        font-size: 18px;
        font-weight: 600;
        color: #2c3e50;
        margin-bottom: 15px;
      }
      .chart-placeholder {
        background: #f8f9fa;
        border: 2px dashed #667eea;
        border-radius: 8px;
        padding: 60px 20px;
        text-align: center;
        color: #6c757d;
      }
      .chart-placeholder i {
        font-size: 48px;
        margin-bottom: 15px;
        color: #667eea;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        background: white;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 3px rgba(0,0,0,0.1);
      }
      th {
        background: #667eea;
        color: white;
        padding: 15px 12px;
        text-align: left;
        font-weight: 600;
        font-size: 14px;
      }
      td {
        padding: 12px;
        border-bottom: 1px solid #e9ecef;
        font-size: 13px;
      }
      tr:nth-child(even) {
        background: #f8f9fa;
      }
      .no-data {
        text-align: center;
        color: #6c757d;
        font-style: italic;
        padding: 40px;
        background: #f8f9fa;
        border-radius: 8px;
        margin: 20px 0;
      }
      .trend-analysis {
        background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%);
        padding: 25px;
        border-radius: 8px;
        margin: 25px 0;
        border-left: 4px solid #ffc107;
      }
      .footer {
        text-align: center;
        color: #6c757d;
        font-size: 12px;
        margin-top: 40px;
        padding-top: 20px;
        border-top: 1px solid #e9ecef;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">
        <h1>KPI Performance Report</h1>
        <h2>${indicator.name}</h2>
      </div>
      
      <div class="info-section">
        <div class="indicator-info">
          <h3>📊 Indicator Information</h3>
          <p><strong>Department:</strong> ${indicator.department}</p>
          <p><strong>Description:</strong> ${indicator.description || 'No description available'}</p>
          <p><strong>Total Data Points:</strong> ${results.length}</p>
        </div>
        
        <div class="targets">
          <h3>🎯 Performance Targets</h3>
          <p><strong>Target:</strong> ${indicator.target ? indicator.target + '%' : 'Not set'}</p>
          <p><strong>Benchmark:</strong> ${indicator.benchmark ? indicator.benchmark + '%' : 'Not set'}</p>
        </div>
      </div>
      
      ${generateStatisticsSection(results)}
      ${generateTrendAnalysis(results)}
      
      <div class="section-title">📈 Performance Chart</div>
      <div class="chart-container">
        <div class="chart-title">Performance Trend Over Time</div>
        <div class="chart-placeholder">
          <i>📊</i>
          <h3>Performance Trend Visualization</h3>
          <p>For interactive charts with detailed visualization, please use the dashboard interface.</p>
          <p><strong>Date Range:</strong> ${results.length > 0 ? formatMonthForDisplay(results[0].month) : 'N/A'} to ${results.length > 0 ? formatMonthForDisplay(results[results.length-1].month) : 'N/A'}</p>
        </div>
      </div>
      
      <div class="section-title">📋 Detailed Results (Chronological Order)</div>
      ${generateResultsTable(results)}
      
      <div class="footer">
        <p>Report generated on ${currentDate}</p>
        <p>KPI Management System - Confidential</p>
      </div>
    </div>
  </body>
  </html>
  `;
  
  return html;
}

function generateStatisticsSection(results) {
  if (results.length === 0) {
    return '<div class="no-data">No data available for statistics</div>';
  }
  
  var numericResults = results.filter(function(r) {
    return !isNaN(parseFloat(r.result));
  }).map(function(r) {
    return parseFloat(r.result);
  });
  
  if (numericResults.length === 0) {
    return '<div class="no-data">No numeric data available for statistics</div>';
  }
  
  var average = numericResults.reduce(function(a, b) { return a + b; }, 0) / numericResults.length;
  var min = Math.min.apply(Math, numericResults);
  var max = Math.max.apply(Math, numericResults);
  var latest = numericResults[numericResults.length - 1];
  
  return `
  <div class="section-title">📊 Performance Statistics</div>
  <div class="stats">
    <div class="stat-box">
      <div class="stat-value">${numericResults.length}</div>
      <div class="stat-label">Data Points</div>
    </div>
    <div class="stat-box">
      <div class="stat-value">${average.toFixed(2)}%</div>
      <div class="stat-label">Average</div>
    </div>
    <div class="stat-box">
      <div class="stat-value">${min.toFixed(2)}%</div>
      <div class="stat-label">Minimum</div>
    </div>
    <div class="stat-box">
      <div class="stat-value">${max.toFixed(2)}%</div>
      <div class="stat-label">Maximum</div>
    </div>
    <div class="stat-box">
      <div class="stat-value">${latest.toFixed(2)}%</div>
      <div class="stat-label">Latest</div>
    </div>
  </div>
  `;
}

function generateTrendAnalysis(results) {
  if (results.length < 2) {
    return '<div class="no-data">Insufficient data for trend analysis (minimum 2 data points required)</div>';
  }
  
  var numericResults = results.filter(function(r) {
    return !isNaN(parseFloat(r.result));
  }).map(function(r) {
    return parseFloat(r.result);
  });
  
  if (numericResults.length < 2) {
    return '<div class="no-data">Insufficient numeric data for trend analysis</div>';
  }
  
  var firstValue = numericResults[0];
  var lastValue = numericResults[numericResults.length - 1];
  var change = lastValue - firstValue;
  var changePercent = firstValue !== 0 ? (change / firstValue) * 100 : 0;
  
  var trend = change > 0 ? 'improving' : change < 0 ? 'declining' : 'stable';
  var trendColor = change > 0 ? '#28a745' : change < 0 ? '#dc3545' : '#6c757d';
  var trendIcon = change > 0 ? '↗' : change < 0 ? '↘' : '→';
  
  return `
  <div class="section-title">📈 Trend Analysis</div>
  <div class="trend-analysis">
    <h3>Performance Trend: <span style="color: ${trendColor}">${trendIcon} ${trend}</span></h3>
    <p><strong>First Value:</strong> ${firstValue.toFixed(2)}%</p>
    <p><strong>Latest Value:</strong> ${lastValue.toFixed(2)}%</p>
    <p><strong>Change:</strong> <span style="color: ${trendColor}; font-weight: bold">${change > 0 ? '+' : ''}${change.toFixed(2)}% (${changePercent.toFixed(1)}%)</span></p>
    <p><strong>Period:</strong> ${results.length} months</p>
    <p><strong>Date Range:</strong> ${formatMonthForDisplay(results[0].month)} to ${formatMonthForDisplay(results[results.length-1].month)}</p>
  </div>
  `;
}

function generateResultsTable(results) {
  if (results.length === 0) {
    return '<div class="no-data">No results available</div>';
  }
  
  var tableHtml = `
  <table>
    <thead>
      <tr>
        <th>Month</th>
        <th>Value</th>
        <th>Numerator</th>
        <th>Denominator</th>
        <th>Result</th>
        <th>Notes</th>
      </tr>
    </thead>
    <tbody>
  `;
  
  results.forEach(function(result) {
    var displayMonth = result.monthFormatted || formatMonthForDisplay(result.month);
    var resultDisplay = '-';
    
    if (result.result) {
      var resultValue = parseFloat(result.result);
      if (!isNaN(resultValue)) {
        if (result.numerator && result.denominator) {
          resultDisplay = resultValue.toFixed(2) + '%';
        } else {
          resultDisplay = resultValue.toFixed(2);
        }
      }
    }
    
    tableHtml += `
      <tr>
        <td>${displayMonth}</td>
        <td>${result.value || '-'}</td>
        <td>${result.numerator || '-'}</td>
        <td>${result.denominator || '-'}</td>
        <td><strong>${resultDisplay}</strong></td>
        <td>${result.note || '-'}</td>
      </tr>
    `;
  });
  
  tableHtml += `
    </tbody>
  </table>
  `;
  
  return tableHtml;
}

// =================================================================
// Export Functions
// =================================================================
function exportIndicatorResultsToExcel(indicatorName) {
  try {
    var indicator = getIndicatorByName(indicatorName);
    if (!indicator) throw new Error('Indicator not found');
    
    var results = getResults(indicator.department, indicator.name);
    
    var sortedResults = results.sort(function(a, b) {
      var dateA = new Date(a.month + '-01');
      var dateB = new Date(b.month + '-01');
      return dateA - dateB;
    });
    
    var spreadsheet = SpreadsheetApp.create('Results_' + indicatorName);
    var sheet = spreadsheet.getActiveSheet();
    
    var headers = ['Month', 'Value', 'Numerator', 'Denominator', 'Result', 'Notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    if (sortedResults.length > 0) {
      var data = sortedResults.map(function(result) {
        var resultDisplay = result.result || '';
        if (result.result && result.numerator && result.denominator) {
          var resultValue = parseFloat(result.result);
          if (!isNaN(resultValue)) {
            resultDisplay = resultValue.toFixed(2) + '%';
          }
        }
        
        return [
          result.monthFormatted || result.month,
          result.value || '',
          result.numerator || '',
          result.denominator || '',
          resultDisplay,
          result.note || ''
        ];
      });
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }
    
    sheet.autoResizeColumns(1, headers.length);
    return spreadsheet.getUrl();
  } catch (error) {
    throw new Error('Export failed: ' + error.message);
  }
}

// =================================================================
// Test Functions
// =================================================================
function testConnection() {
  try {
    var indicators = getAllIndicators();
    var users = getCurrentUserRole();
    var departments = getDepartments();
    
    return {
      status: 'success',
      indicatorsCount: indicators.length,
      departmentsCount: departments.length,
      departments: departments,
      currentUser: users
    };
  } catch (error) {
    return {
      status: 'error',
      error: error.message
    };
  }
}
