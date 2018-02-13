function TwoWeekSprintCopyData() {
  var CHART_TITLES = ['Adaptive Platform','One Click'];
  var CHART_RANGES = ['A1:B8','A13:B20','A25:B32','G25:H32','A37:C45','H1:I8'];
  var VELOCITY_CHART = 'Velocity Charts'
  
  function moveDataByLine() {
    var velocitySheet = getSheet(VELOCITY_CHART);
    velocitySheet.getRange('A3:B8').moveTo(velocitySheet.getRange('A2:B7'));
    velocitySheet.getRange('A15:B20').moveTo(velocitySheet.getRange('A14:B19'));
  }
  
  function setChartsDataRange() {
    var velocitySheet = getSheet(VELOCITY_CHART);
    var charts = velocitySheet.getCharts();    
    charts.forEach(removeDataRangeAndAddNew);
  }

  function removeDataRangeAndAddNew(chart, index) {
    var chartTitle = chart.getOptions().get('title');
    if (CHART_TITLES.indexOf(chartTitle) < 0)
      return;
    
    var velocitySheet = getSheet(VELOCITY_CHART);
    var ranges = chart.getRanges();
    var chartBuilder = chart.modify();
    
    ranges.forEach(function removeRagne(range) {
      chartBuilder.removeRange(range);
    });
    
    var rangeToAdd = velocitySheet.getRange(CHART_RANGES[index]);
    chartBuilder.addRange(rangeToAdd);
    velocitySheet.updateChart(chartBuilder.build());
  }
  
  function copyVelocities() {
    var apValue = getRangeValues('A8:B8');
    var oneClickValue = getRangeValues('A20:B20');

    writeVauesToHistoric('Adaptive Platform', apValue);
    writeVauesToHistoric('One Click', oneClickValue);
  }
  
  function getRangeValues(range) {
    var velocitySheet = getSheet(VELOCITY_CHART);
    var lastWeekVelocityRange = velocitySheet.getRange(range).getValues();
    return lastWeekVelocityRange[0]; //Get the first line (assumption we're copying single line only) 
  }
  
  function writeVauesToHistoric(historicSheetName, values) {
    var historicSheet = getSheet(historicSheetName);
    historicSheet.appendRow(values);
  }
  
  function getSheet(name) {
    return sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  }
  
  copyVelocities();
  moveDataByLine();
  setChartsDataRange();

}