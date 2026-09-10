using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Analizator.Core;

internal static class StatisticsChartWriter
{
    private const string ChartUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public static void AddCharts(
        string workbookPath,
        IReadOnlyList<string> monthLabels,
        IReadOnlyList<double> monthlyAttempts,
        IReadOnlyList<double> monthlyPassed,
        IReadOnlyList<string> installationLabels,
        IReadOnlyList<double> installationAttempts)
    {
        if (monthLabels.Count <= 0 || installationLabels.Count <= 0)
            return;
        if (monthlyAttempts.Count != monthLabels.Count || monthlyPassed.Count != monthLabels.Count ||
            installationAttempts.Count != installationLabels.Count)
            throw new ArgumentException("Количество подписей диаграмм не совпадает с количеством значений.");

        using var document = SpreadsheetDocument.Open(workbookPath, true);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidOperationException("В книге отсутствует основная часть Excel.");
        var summaryPart = GetWorksheetPart(workbookPart, "Сводка");
        var drawingsPart = summaryPart.DrawingsPart ?? summaryPart.AddNewPart<DrawingsPart>();
        if (summaryPart.Worksheet.Elements<S.Drawing>().FirstOrDefault() is null)
        {
            summaryPart.Worksheet.Append(new S.Drawing
            {
                Id = summaryPart.GetIdOfPart(drawingsPart)
            });
        }
        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

        var displayedMonthLabels = monthLabels.TakeLast(24).ToArray();
        var displayedMonthlyAttempts = monthlyAttempts.TakeLast(24).ToArray();
        var displayedMonthlyPassed = monthlyPassed.TakeLast(24).ToArray();
        var monthlyLastRow = 6 + monthLabels.Count;
        var monthlyFirstRow = Math.Max(7, monthlyLastRow - 23);
        AddLineChart(
            drawingsPart,
            "Динамика прохождений",
            SheetFormula("Динамика", $"$A${monthlyFirstRow}:$A${monthlyLastRow}"),
            displayedMonthLabels,
            [
                new ChartSeries(
                    "Попытки",
                    SheetFormula("Динамика", "$B$6"),
                    SheetFormula("Динамика", $"$B${monthlyFirstRow}:$B${monthlyLastRow}"),
                    "5B9BD5",
                    displayedMonthlyAttempts),
                new ChartSeries(
                    "Успешно",
                    SheetFormula("Динамика", "$F$6"),
                    SheetFormula("Динамика", $"$F${monthlyFirstRow}:$F${monthlyLastRow}"),
                    "70AD47",
                    displayedMonthlyPassed)
            ],
            fromColumn: 10,
            fromRow: 5,
            toColumn: 20,
            toRow: 20,
            axisSeed: 48_650_100U);

        var displayedInstallationLabels = installationLabels.Take(12).ToArray();
        var displayedInstallationAttempts = installationAttempts.Take(12).ToArray();
        var installationLastRow = 6 + displayedInstallationLabels.Length;
        AddBarChart(
            drawingsPart,
            "Количество попыток по установкам",
            SheetFormula("Установки", $"$A$7:$A${installationLastRow}"),
            displayedInstallationLabels,
            new ChartSeries(
                "Попытки",
                SheetFormula("Установки", "$D$6"),
                SheetFormula("Установки", $"$D$7:$D${installationLastRow}"),
                "4472C4",
                displayedInstallationAttempts),
            fromColumn: 10,
            fromRow: 21,
            toColumn: 20,
            toRow: 37,
            axisSeed: 48_650_200U);

        drawingsPart.WorksheetDrawing.Save();
        summaryPart.Worksheet.Save();
    }

    private static void AddLineChart(
        DrawingsPart drawingsPart,
        string title,
        string categoriesFormula,
        IReadOnlyList<string> categories,
        IReadOnlyList<ChartSeries> series,
        int fromColumn,
        int fromRow,
        int toColumn,
        int toRow,
        uint axisSeed)
    {
        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        var chart = CreateChart(title);
        var plotArea = chart.GetFirstChild<C.PlotArea>()!;
        var lineChart = new C.LineChart(
            new C.Grouping { Val = C.GroupingValues.Standard },
            new C.VaryColors { Val = false });

        for (var index = 0; index < series.Count; index++)
        {
            var item = series[index];
            var chartSeries = new C.LineChartSeries(
                new C.Index { Val = (uint)index },
                new C.Order { Val = (uint)index },
                CreateSeriesText(item.NameFormula, item.Name),
                new C.ChartShapeProperties(CreateOutline(item.Color)),
                new C.Marker(
                    new C.Symbol { Val = C.MarkerStyleValues.Circle },
                    new C.Size { Val = 6 }),
                new C.CategoryAxisData(CreateStringReference(categoriesFormula, categories)),
                new C.Values(CreateNumberReference(item.ValuesFormula, item.Values)));
            lineChart.Append(chartSeries);
        }

        lineChart.Append(new C.AxisId { Val = axisSeed });
        lineChart.Append(new C.AxisId { Val = axisSeed + 1 });
        plotArea.Append(lineChart);
        AppendAxes(plotArea, axisSeed, axisSeed + 1, C.AxisPositionValues.Bottom,
            C.AxisPositionValues.Left, "#,##0");
        FinalizeChart(chart, showLegend: true);
        chartPart.ChartSpace = WrapChart(chart);
        AddAnchor(drawingsPart, chartPart, title, fromColumn, fromRow, toColumn, toRow);
    }

    private static void AddBarChart(
        DrawingsPart drawingsPart,
        string title,
        string categoriesFormula,
        IReadOnlyList<string> categories,
        ChartSeries series,
        int fromColumn,
        int fromRow,
        int toColumn,
        int toRow,
        uint axisSeed)
    {
        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        var chart = CreateChart(title);
        var plotArea = chart.GetFirstChild<C.PlotArea>()!;
        var barChart = new C.BarChart(
            new C.BarDirection { Val = C.BarDirectionValues.Bar },
            new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
            new C.VaryColors { Val = false },
            new C.BarChartSeries(
                new C.Index { Val = 0U },
                new C.Order { Val = 0U },
                CreateSeriesText(series.NameFormula, series.Name),
                new C.ChartShapeProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = series.Color }),
                    new A.Outline(new A.NoFill())),
                new C.CategoryAxisData(CreateStringReference(categoriesFormula, categories)),
                new C.Values(CreateNumberReference(series.ValuesFormula, series.Values))),
            new C.GapWidth { Val = 70 });
        barChart.Append(new C.AxisId { Val = axisSeed });
        barChart.Append(new C.AxisId { Val = axisSeed + 1 });
        plotArea.Append(barChart);
        AppendAxes(plotArea, axisSeed, axisSeed + 1, C.AxisPositionValues.Left,
            C.AxisPositionValues.Bottom, "#,##0");
        FinalizeChart(chart, showLegend: false);
        chartPart.ChartSpace = WrapChart(chart);
        AddAnchor(drawingsPart, chartPart, title, fromColumn, fromRow, toColumn, toRow);
    }

    private static C.Chart CreateChart(string title)
    {
        var chart = new C.Chart(
            new C.AutoTitleDeleted { Val = false },
            CreateTitle(title),
            new C.PlotArea(new C.Layout()));
        return chart;
    }

    private static C.ChartSpace WrapChart(C.Chart chart) =>
        new(
            new C.EditingLanguage { Val = "ru-RU" },
            new C.RoundedCorners { Val = false },
            chart);

    private static void FinalizeChart(C.Chart chart, bool showLegend)
    {
        if (showLegend)
        {
            chart.Append(new C.Legend(
                new C.LegendPosition { Val = C.LegendPositionValues.Top },
                new C.Layout(),
                new C.Overlay { Val = false }));
        }
        chart.Append(new C.PlotVisibleOnly { Val = true });
        chart.Append(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });
    }

    private static void AppendAxes(
        C.PlotArea plotArea,
        uint categoryAxisId,
        uint valueAxisId,
        C.AxisPositionValues categoryPosition,
        C.AxisPositionValues valuePosition,
        string numberFormat)
    {
        plotArea.Append(new C.CategoryAxis(
            new C.AxisId { Val = categoryAxisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = categoryPosition },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = valueAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true },
            new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset { Val = 100 }));

        plotArea.Append(new C.ValueAxis(
            new C.AxisId { Val = valueAxisId },
            new C.Scaling(
                new C.Orientation { Val = C.OrientationValues.MinMax },
                new C.MinAxisValue { Val = 0D }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = valuePosition },
            new C.MajorGridlines(),
            new C.NumberingFormat { FormatCode = numberFormat, SourceLinked = false },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = categoryAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between }));
    }

    private static C.Title CreateTitle(string title) =>
        new(
            new C.ChartText(
                new C.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties { Language = "ru-RU", FontSize = 1200 },
                            new A.Text(title))))),
            new C.Overlay { Val = false });

    private static A.Outline CreateOutline(string color)
    {
        var outline = new A.Outline { Width = 25_400 };
        outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
        return outline;
    }

    private static C.SeriesText CreateSeriesText(string formula, string value) =>
        new(CreateStringReference(formula, [value]));

    private static C.StringReference CreateStringReference(
        string formula,
        IReadOnlyList<string> values)
    {
        var cache = new C.StringCache(new C.PointCount { Val = (uint)values.Count });
        for (var index = 0; index < values.Count; index++)
        {
            cache.Append(new C.StringPoint(
                new C.NumericValue(values[index]))
            {
                Index = (uint)index
            });
        }
        return new C.StringReference(new C.Formula(formula), cache);
    }

    private static C.NumberReference CreateNumberReference(
        string formula,
        IReadOnlyList<double> values)
    {
        var cache = new C.NumberingCache(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)values.Count });
        for (var index = 0; index < values.Count; index++)
        {
            cache.Append(new C.NumericPoint(
                new C.NumericValue(values[index].ToString(CultureInfo.InvariantCulture)))
            {
                Index = (uint)index
            });
        }
        return new C.NumberReference(new C.Formula(formula), cache);
    }

    private static void AddAnchor(
        DrawingsPart drawingsPart,
        ChartPart chartPart,
        string name,
        int fromColumn,
        int fromRow,
        int toColumn,
        int toRow)
    {
        var drawingId = drawingsPart.WorksheetDrawing!
            .Descendants<Xdr.NonVisualDrawingProperties>()
            .Select(item => item.Id?.Value ?? 0U)
            .DefaultIfEmpty(0U).Max() + 1U;
        var chartReference = new C.ChartReference
        {
            Id = drawingsPart.GetIdOfPart(chartPart)
        };
        var graphicFrame = new Xdr.GraphicFrame(
            new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties { Id = drawingId, Name = name },
                new Xdr.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoGrouping = true })),
            new Xdr.Transform(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L }),
            new A.Graphic(
                new A.GraphicData(chartReference) { Uri = ChartUri }));

        drawingsPart.WorksheetDrawing!.Append(new Xdr.TwoCellAnchor(
            Marker<Xdr.FromMarker>(fromColumn, fromRow),
            Marker<Xdr.ToMarker>(toColumn, toRow),
            graphicFrame,
            new Xdr.ClientData()));
    }

    private static T Marker<T>(int column, int row) where T : OpenXmlCompositeElement, new()
    {
        var marker = new T();
        marker.Append(
            new Xdr.ColumnId((column - 1).ToString(CultureInfo.InvariantCulture)),
            new Xdr.ColumnOffset("0"),
            new Xdr.RowId((row - 1).ToString(CultureInfo.InvariantCulture)),
            new Xdr.RowOffset("0"));
        return marker;
    }

    private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
    {
        var sheet = workbookPart.Workbook.Sheets?.Elements<S.Sheet>()
            .FirstOrDefault(item => string.Equals(item.Name?.Value, sheetName,
                StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"В книге отсутствует лист «{sheetName}».");
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
    }

    private static string SheetFormula(string sheet, string range) =>
        $"'{sheet.Replace("'", "''")}'!{range}";

    private sealed record ChartSeries(
        string Name,
        string NameFormula,
        string ValuesFormula,
        string Color,
        IReadOnlyList<double> Values);
}
