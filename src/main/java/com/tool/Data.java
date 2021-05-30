package com.tool;

public class Data {
    double displacement;
    double pressure;
    public Data(double displacement, double pressure) {
        this.displacement = displacement;
        this.pressure = pressure;
    }

    @Override
    public String toString() {
        return "Data{" +
                "displacement=" + displacement +
                ", pressure=" + pressure +
                '}';
    }
}
