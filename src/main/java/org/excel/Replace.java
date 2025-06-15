package org.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import org.apache.poi.util.StringUtil;

import java.lang.reflect.Array;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

public class Replace {
    //    public static final String sourceFile = "src/main/resources/学生名单.xls";
    public static final String sourceFile = "学生名单.xls";
    public static final String targetFile = "预录取名单.xls";
    //    public static final String targetFile = "src/main/resources/预录取名单.xls";
    public static final String sourceSheetName = "学生名单";
    public static final String targetSheetName = "预录取名单";
    public static final List<String> sheetKeys = Arrays.asList("中考报名号", "证件号码", "姓名");

    public static void main(String[] args) {
        LinkedHashMap sourceHeaders = new LinkedHashMap();
        List<LinkedHashMap> sourceFiles = readFiles(sourceFile, sourceHeaders, sourceSheetName);
        LinkedHashMap<Object, Object> targetHeaders = new LinkedHashMap<>();
        List<LinkedHashMap> targetFiles = readFiles(targetFile, targetHeaders, targetSheetName);

        Map<String, Integer> sheetKeyIndex = new HashMap<>();
        Map<Integer, String> sheetKeyIndexX = new HashMap<>();

        sourceHeaders.forEach((k, vv) -> {
            int indexOf = sheetKeys.indexOf((String) vv);
            if (indexOf >= 0) {
                sheetKeyIndex.put((String) vv, (Integer) k);
                sheetKeyIndexX.put((Integer) k, (String) vv);
            }
        });

        for (int i = 0; i < sourceFiles.size(); i++) {
            LinkedHashMap v = sourceFiles.get(i);
            String[] datas = new String[sheetKeys.size()];
            final AtomicInteger count = new AtomicInteger(0);
            v.forEach((k, vv) -> {
                Integer kk = (Integer) k;
                String vvv = (String) vv;
                if (vvv.trim().length() == 0) {
                    return;
                }

                if (!sheetKeyIndexX.keySet().contains(kk)) {
                    return;
                }

                String s = sheetKeyIndexX.get(kk);
                int indexOf1 = sheetKeys.indexOf(s);
                datas[indexOf1] = vvv;
                count.addAndGet(1);
            });

            if (count.get() == sheetKeys.size()) {
                LinkedHashMap datasx = new LinkedHashMap();
                for (int ii = 0; ii < datas.length; ii++) {
                    datasx.put(ii, datas[ii]);
                }
                targetFiles.add(datasx);
            }
        }

        Optional<Object> max = targetHeaders.keySet().stream().max(Comparator.comparingInt(a -> (Integer) a));
        final List<List<String>> headers = new ArrayList<>();
        for (int i = 0; i < (Integer) max.get() + 1; i++) {
            headers.add(Arrays.asList());
        }
        targetHeaders.forEach((k, v) -> {
            headers.set((Integer) k, Arrays.asList((String) v));
        });
        EasyExcel.write(targetFile, LinkedHashMap.class).sheet(targetSheetName).head(headers).doWrite(targetFiles);
    }

    private static List<LinkedHashMap> readFiles(String name, LinkedHashMap headers, String sheetName) {
        CommonListener commonListener = new CommonListener();
        commonListener.setSheetName(sheetName);
        ExcelReaderBuilder builder = EasyExcel.read(name, commonListener);
        builder.doReadAll();
        headers.putAll(commonListener.getHead());
        return commonListener.getCachedDataList();
    }
}