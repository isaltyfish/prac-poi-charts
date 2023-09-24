package net.verytools.prac.model;

import org.apache.poi.xddf.usermodel.PresetColor;

import java.util.List;

public class SerieData {
    /**
     * 系列名称
     */
    private String name;
    /**
     * 系列数据
     */
    private List<Number> data;
    /**
     * 系列的颜色
     */
    private PresetColor color = PresetColor.ALICE_BLUE;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Number> getData() {
        return data;
    }

    public void setData(List<Number> data) {
        this.data = data;
    }

    public PresetColor getColor() {
        return color;
    }

    public void setColor(PresetColor color) {
        this.color = color;
    }
}
