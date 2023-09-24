package net.verytools.prac;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileOutputStream;
import java.io.IOException;

public class LineChartRenderer {

    public void render(XWPFDocument doc) throws IOException, InvalidFormatException {
        // 创建图
        XWPFChart chart = doc.createChart(14 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);
        // 标题
        chart.setTitleText("专业课选修人数");
        chart.setTitleOverlay(false);
        // 图例位置
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        // X轴标题位置
        XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        xAxis.setTitle("专业名称");
        // y轴标题位置
        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
        yAxis.setTitle("数量");

        // 创建分类数据源和值数据源
        XDDFCategoryDataSource categoryDS = XDDFDataSourcesFactory.fromArray(new String[]{"英语", "数学", "计算机", "财经"});
        XDDFNumericalDataSource<Integer> valueDs = XDDFDataSourcesFactory.fromArray(new Integer[]{100, 200, 800, 50});
        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, xAxis, yAxis);
        // 将数据源绑定到图表中
        XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(categoryDS, valueDs);

        // 折线图例标题
        series1.setTitle("选修人数", null);
        // 直线
        series1.setSmooth(false);
        // 设置标记大小
        series1.setMarkerSize((short) 6);
        // 设置标记样式，星星
        series1.setMarkerStyle(MarkerStyle.STAR);
        chart.plot(data);
        // Write the output to a file
        try (FileOutputStream fo = new FileOutputStream("D:\\line.docx")) {
            doc.write(fo);
        }
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        new LineChartRenderer().render(new XWPFDocument());
    }

}
