package org.excel;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.util.ListUtils;
import lombok.Data;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Data
public class CommonListener implements ReadListener<LinkedHashMap> {

    private static final int BATCH_COUNT = 100;

    private List<LinkedHashMap> cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);

    private String sheetName;

    private LinkedHashMap head = new LinkedHashMap();


    @Override
    public void invoke(LinkedHashMap data, AnalysisContext context) {
        String sheetName = (String) context.readSheetHolder().getSheetName();
        if (!sheetName.trim().equals(this.sheetName)) {
            return;
        }
        cachedDataList.add(data);
    }

    @Override
    public void invokeHead(Map<Integer, ReadCellData<?>> headMap, AnalysisContext context) {
        String sheetName = (String) context.readSheetHolder().getSheetName();
        if (!sheetName.trim().equals(this.sheetName)) {
            return;
        }

        headMap.forEach((k, v) -> {
            this.head.put(k, v.getStringValue());
        });

    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
    }
}
