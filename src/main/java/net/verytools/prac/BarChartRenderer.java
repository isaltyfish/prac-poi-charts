package net.verytools.prac;

import net.verytools.prac.model.ChartData;
import net.verytools.prac.model.SerieData;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 柱状图绘制
 */
public class BarChartRenderer {

    public void render(XWPFDocument document, ChartData chartData) throws IOException, InvalidFormatException {
        // create the chart
        XWPFChart chart = document.createChart(14 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);
        // create axis
        XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
        yAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
        // Else first and last category is exactly on cross points and the bars are only half visible.
        yAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

        // create data sources
        String[] categories = chartData.getCategories().toArray(new String[0]);
        XDDFDataSource<String> categoryDS = XDDFDataSourcesFactory.fromArray(categories);
        List<XDDFNumericalDataSource<Number>> valueDS = new ArrayList<>();
        List<SerieData> seriesData = chartData.getSeries();
        for (SerieData seriesDatum : seriesData) {
            XDDFNumericalDataSource<Number> s = XDDFDataSourcesFactory.fromArray(seriesDatum.getData().toArray(new Number[0]));
            valueDS.add(s);
        }

        // create chart data
        XDDFChartData data = chart.createData(ChartTypes.BAR, xAxis, yAxis);
        ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);
        ((XDDFBarChartData) data).setGapWidth(500);

        int i = 0;
        data.setVaryColors(seriesData.size() > 1);
        for (XDDFNumericalDataSource<Number> value : valueDS) {
            XDDFChartData.Series series = data.addSeries(categoryDS, value);
            series.setTitle(seriesData.get(i).getName(), null);
            solidFillSeries(series, seriesData.get(i).getColor());
            i++;
        }

        // 图例
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);
        legend.setOverlay(false);

        chart.plot(data);
    }

    private static void solidFillSeries(XDDFChartData.Series series, PresetColor color) {
        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(color));
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series.setShapeProperties(properties);
    }

    public static void main(String[] args) throws Exception {
        ChartData chartData = new ChartData();
        chartData.setTitle("专业课选修人数/课时数");
        chartData.setCategories(Arrays.asList("英语", "数学", "计算机", "财经"));

        // 构造系列数据
        List<SerieData> seriesData = new ArrayList<>();
        seriesData.add(createSerieData("选修人数", Arrays.asList(100, 200, 800, 50), PresetColor.RED));
        seriesData.add(createSerieData("课时数", Arrays.asList(50, 60, 40, 75), PresetColor.DARK_GREEN));
        chartData.setSeries(seriesData);

        XWPFDocument document = new XWPFDocument();
        new BarChartRenderer().render(document, chartData);
        try (FileOutputStream fileOut = new FileOutputStream("D:\\bar.docx")) {
            document.write(fileOut);
        }
    }

    private static SerieData createSerieData(String name, List<Number> data, PresetColor color) {
        SerieData ret = new SerieData();
        ret.setName(name);
        ret.setColor(color);
        ret.setData(data);
        return ret;
    }

}
