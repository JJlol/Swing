package com.swing.bean;

import java.util.List;

public class Data {
    private String name;
    private List<List<Object>> value;
    private Double upperLimit;
    private Double lowLimit;
    private Double max = Double.MIN_VALUE;
    private Double min = Double.MAX_VALUE;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<List<Object>> getValue() {
        return value;
    }

    public void setValue(List<List<Object>> value) {
        this.value = value;
    }

    public Double getUpperLimit() {
        return upperLimit;
    }

    public void setUpperLimit(Double upperLimit) {
        this.upperLimit = upperLimit;
    }

    public Double getLowLimit() {
        return lowLimit;
    }

    public void setLowLimit(Double lowLimit) {
        this.lowLimit = lowLimit;
    }

    public Double getMax() {
        return max;
    }

    public void setMax(Double max) {
        this.max = max;
    }

    public Double getMin() {
        return min;
    }

    public void setMin(Double min) {
        this.min = min;
    }
}
