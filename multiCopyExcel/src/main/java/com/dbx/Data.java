package com.dbx;

import com.alibaba.excel.annotation.ExcelProperty;

import java.util.Objects;

public class Data {
    @ExcelProperty(index = 0)
    private String code;
    @ExcelProperty(index = 3)
    private String firstPit;

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getFirstPit() {
        return firstPit;
    }

    public void setFirstPit(String firstPit) {
        this.firstPit = firstPit;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Data data = (Data) o;
        return Objects.equals(code, data.code) &&
                Objects.equals(firstPit, data.firstPit);
    }

    @Override
    public int hashCode() {
        return Objects.hash(code, firstPit);
    }

    @Override
    public String toString() {
        return "Data{" +
                "code='" + code + '\'' +
                ", firstPit='" + firstPit + '\'' +
                '}';
    }
}
