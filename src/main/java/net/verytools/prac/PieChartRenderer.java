package net.verytools.prac;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;

import java.io.FileOutputStream;
import java.io.IOException;

public class PieChartRenderer {

    public void render() throws IOException, InvalidFormatException {
        XWPFDocument document = new XWPFDocument();
        // create the chart
        XWPFChart chart = document.createChart(14 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);
        // 标题
        chart.setTitleText("专业课选修人数");
        // 标题是否覆盖图表
        chart.setTitleOverlay(false);

        // 图例位置
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);

        // 分类数据源
        XDDFCategoryDataSource categoryDS = XDDFDataSourcesFactory.fromArray(new String[]{"英语", "数学", "计算机", "财经"});
        // 值数据源
        XDDFNumericalDataSource<Number> valuesDS = XDDFDataSourcesFactory.fromArray(new Integer[]{10, 20, 30, 50});
        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        // 设置为可变颜色
        data.setVaryColors(true);
        // 绑定数据源到图表
        XDDFChartData.Series series = data.addSeries(categoryDS, valuesDS);

        series.setShowLeaderLines(true);
        // 隐藏图例标识、系列名称、分类名称和数值
        XDDFPieChartData.Series s = (XDDFPieChartData.Series) series;
        CTPieSer ctPieSer = s.getCTPieSer();
        showCateName(ctPieSer, false);
        showVal(ctPieSer, false);
        showLegendKey(ctPieSer, false);
        showSerName(ctPieSer, false);
        // 绘制
        chart.plot(data);

//        try (FileOutputStream fileOut = new FileOutputStream("D:\\pie.docx")) {
        try (FileOutputStream fileOut = new FileOutputStream("/tmp/pie.docx")) {
            document.write(fileOut);
        }
    }

    public void showCateName(CTPieSer series, boolean val) {
        if (series.getDLbls().isSetShowCatName()) {
            series.getDLbls().getShowCatName().setVal(val);
        } else {
            series.getDLbls().addNewShowCatName().setVal(val);
        }
    }

    public void showVal(CTPieSer series, boolean val) {
        if (series.getDLbls().isSetShowVal()) {
            series.getDLbls().getShowVal().setVal(val);
        } else {
            series.getDLbls().addNewShowVal().setVal(val);
        }
    }

    public void showSerName(CTPieSer series, boolean val) {
        if (series.getDLbls().isSetShowSerName()) {
            series.getDLbls().getShowSerName().setVal(val);
        } else {
            series.getDLbls().addNewShowSerName().setVal(val);
        }
    }

    public void showLegendKey(CTPieSer series, boolean val) {
        if (series.getDLbls().isSetShowLegendKey()) {
            series.getDLbls().getShowLegendKey().setVal(val);
        } else {
            series.getDLbls().addNewShowLegendKey().setVal(val);
        }
    }

    public static void main(String[] args) throws Exception {
        new PieChartRenderer().render();
    }

}
