function OneWeekSprintCopyData() {
  var CHART_TITLES = ['Future Funding','IT Ops', 'IT Ops - total closed tickets', 'UCM Open and UCM Closed'];
  var CHART_RANGES = ['A1:B8','A13:B20','A25:B32','G25:H32','A37:C45','H1:I8'];
  var VELOCITY_CHART = 'Velocity Charts'
  
  function moveDataByWeek() {
    var velocitySheet = getSheet(VELOCITY_CHART);
    velocitySheet.getRange('H3:I8').moveTo(velocitySheet.getRange('H2:I7'));
    velocitySheet.getRange('A27:B32').moveTo(velocitySheet.getRange('A26:B31'));
    velocitySheet.getRange('G27:H32').moveTo(velocitySheet.getRange('G26:H31'));
    velocitySheet.getRange('A39:C45').moveTo(velocitySheet.getRange('A38:C44'));
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
    var futureFundingValue = getRangeValues('H8:I8');
    var itOpsValues = getRangeValues('A32:B32');
    itOpsValues = itOpsValues.concat(getRangeValues('H32:H32'));
    itOpsValues = itOpsValues.concat(getRangeValues('B45:C45'));

    writeVauesToHistoric('Future Funding', futureFundingValue);
    writeVauesToHistoric('IT Ops', itOpsValues);
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
  moveDataByWeek();
  setChartsDataRange();

}
