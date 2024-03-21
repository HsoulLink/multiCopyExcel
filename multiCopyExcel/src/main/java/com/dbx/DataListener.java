package com.dbx;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.util.ListUtils;

import java.util.List;

public class DataListener implements ReadListener<Data> {

    private static final int BATCH_COUNT = 100;

    private List<Data> dataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);

    public List<Data> getDataList() {
        return dataList;
    }

    @Override
    public void invoke(Data data, AnalysisContext analysisContext) {
        dataList.add(data);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println("读取完成!");
    }
}
