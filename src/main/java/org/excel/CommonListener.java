package org.excel;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.read.metadata.holder.ReadSheetHolder;
import com.alibaba.excel.util.ListUtils;
import lombok.Data;
import lombok.Getter;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Data
public class CommonListener implements ReadListener<LinkedHashMap> {
    /**
     * 每隔5条存储数据库，实际使用中可以100条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 100;
    /**
     * 缓存的数据
     */
    private List<LinkedHashMap> cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);

    private String sheetName;

    private LinkedHashMap head = new LinkedHashMap();

    /**
     * 这个每一条数据解析都会来调用
     *
     * @param data    one row value. It is same as {@link AnalysisContext#readRowHolder()}
     * @param context
     */
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

    /**
     * 所有数据解析完成了 都会来调用
     *
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
    }
}
