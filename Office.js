
async function buildFullDashboard() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const dashboard = sheets.getItem("Dashboard");

    //  MILESTONE 1: Calculation of revenue and weighted margins
    const revRange = dashboard.getRange("C8:C39");
    revRange.formulas = [["=SUMIFS('Raw Data'!$E$2:$E$187,'Raw Data'!$D$2:$D$187,$A8,'Raw Data'!$B$2:$B$187,VALUE(LEFT($B8,4)),'Raw Data'!$C$2:$C$187,RIGHT($B8,2))"]];
    revRange.numberFormat = [["$#,##0"]];

    const marginRange = dashboard.getRange("D8:D39");
    marginRange.formulas = [["=SUMPRODUCT(('Raw Data'!$D$2:$D$187=$A8)*('Raw Data'!$B$2:$B$187=VALUE(LEFT($B8,4)))*('Raw Data'!$C$2:$C$187=RIGHT($B8,2))*'Raw Data'!$G$2:$G$187)/C8"]];
    marginRange.numberFormat = [["0.0%"]];

    //  MILESTONE 2: Adding trend and comparison products
    const trendRange = dashboard.getRange("E8:E39");
    trendRange.formulas = [["=IF(AND(LEFT($B8,4)=\"2023\",RIGHT($B8,2)=\"Q1\"),\"N/A\",D8-D7)"]];
    trendRange.numberFormat = [["0.0%"]];

    const yoyRange = dashboard.getRange("F8:F39");
    yoyRange.formulas = [["=IF(LEFT($B8,4)=\"2023\",\"N/A\",D8-INDEX($D$8:$D$39,MATCH($A8&(LEFT($B8,4)-1)&RIGHT($B8,2),$A$8:$A$39&LEFT($B$8:$B$39,4)&RIGHT($B$8:$B$39,2),0)))"]];
    yoyRange.numberFormat = [["0.0%"]];

    const healthRange = dashboard.getRange("G8:G39");
    healthRange.formulas = [["=IF(D8>0.35,\"Strong\",IF(D8>=0.2,\"Moderate\",\"At Risk\"))"]];

    
    const cfStrong = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
    cfStrong.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "Strong" };
    cfStrong.textComparison.format.fill.color = "green"; 

    const cfModerate = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
    cfModerate.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "Moderate" };
    cfModerate.textComparison.format.fill.color = "yellow"; 

    const cfAtRisk = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
    cfAtRisk.textComparison.rule = { operator: Excel.ConditionalTextOperator.contains, text: "At Risk" };
    cfAtRisk.textComparison.format.fill.color = "red"; 

    //  MILESTONE 3: Creating the chart
    dashboard.getRange("B44:B51").formulas = [["=SUMPRODUCT(($A$8:$A$39=\"Widget Pro\")*($B$8:$B$39=$A44)*($D$8:$D$39))"]];
    dashboard.getRange("C44:C51").formulas = [["=SUMPRODUCT(($A$8:$A$39=\"Widget Standard\")*($B$8:$B$39=$A44)*($D$8:$D$39))"]];
    dashboard.getRange("D44:D51").formulas = [["=SUMPRODUCT(($A$8:$A$39=\"Service Package\")*($B$8:$B$39=$A44)*($D$8:$D$39))"]];
    dashboard.getRange("E44:E51").formulas = [["=SUMPRODUCT(($A$8:$A$39=\"Accessory Kit\")*($B$8:$B$39=$A44)*($D$8:$D$39))"]];
    dashboard.getRange("F44:F51").formulas = [["=SUMIF($B$8:$B$39,$A44,$C$8:$C$39)"]];

    const chartData = dashboard.getRange("A43:F51");
    const chart = dashboard.charts.add(Excel.ChartType.columnClustered, chartData, Excel.ChartSeriesBy.columns);
    
    chart.title.text = "Quarterly Margin Trends by Product";
    chart.axes.valueAxis.title.text = "Profit Margin";
    
    const revSeries = chart.series.getItemAt(4);
    revSeries.axisGroup = Excel.ChartAxisGroup.secondary;
    revSeries.chartType = Excel.ChartType.line;
    
    chart.axes.secondaryValueAxis.title.text = "Total Revenue ($)";
    chart.axes.secondaryValueAxis.title.visible = true;

    await context.sync();
  });
}