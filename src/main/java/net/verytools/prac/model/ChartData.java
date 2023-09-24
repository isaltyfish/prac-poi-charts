package net.verytools.prac.model;

import java.util.List;

public class ChartData {
    /**
     * 图表的标题
     */
    private String title;
    /**
     * 分类
     */
    private List<String> categories;
    /**
     * 每个系列的数据
     */
    private List<SerieData> series;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public List<String> getCategories() {
        return categories;
    }

    public void setCategories(List<String> categories) {
        this.categories = categories;
    }

    public List<SerieData> getSeries() {
        return series;
    }

    public void setSeries(List<SerieData> series) {
        this.series = series;
    }
}
