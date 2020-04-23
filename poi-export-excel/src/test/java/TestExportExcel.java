import export.entity.TableHeader;
import export.excel.ExportExcel;
import export.excel.NomalExportExcel;
import org.apache.poi.ss.usermodel.Font;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TestExportExcel {

    @Test
    public void test02() throws IOException {
        List<TableHeader> tableHeaderList = new ArrayList<>();
        for (int j = 0; j < 20; j++) {
            TableHeader tableHeader = new TableHeader();
            tableHeader.setHeaderText("第" + j + "列");
            tableHeader.setAlign("center");
            tableHeaderList.add(tableHeader);
        }

        List<List<String>> tableData = new ArrayList<>();
        for (int i = 0; i < 54812; i++) {
            List<String> rowData = new ArrayList<>();
            for (int j = 0; j < 20; j++) {
                rowData.add("单元格" + i + "," + j);
            }
            tableData.add(rowData);
        }

        NomalExportExcel nomalExportExcel = new NomalExportExcel(tableHeaderList, tableData);

        nomalExportExcel.export("D:\\test", "test");
    }

    @Test
    public void test01() throws IOException {
        ExportExcel exportExcel = new ExportExcel();
        exportExcel.createExcel("test1");

        List<TableHeader> tableHeaderList = new ArrayList<>();
        TableHeader tableHeader0 = new TableHeader();
        tableHeader0.setHeaderText("第_1列");
        tableHeader0.setField("field_1.data[0].field_1_1_1.name");
        tableHeader0.setWidth(10);
        tableHeader0.setWrapText(true);
        tableHeader0.setBackground("#4394ff");
        tableHeader0.setAlign("center");
        tableHeaderList.add(tableHeader0);

        for (int i = 0; i < 1; i++) {
            TableHeader tableHeader = new TableHeader();
            tableHeader.setHeaderText("第" + i + "列");
            tableHeader.setField("field" + i);

            List<TableHeader> children = new ArrayList<>();
            for (int j = 0; j < 3; j++) {
                TableHeader tableHeader11 = new TableHeader();
                tableHeader11.setHeaderText("第" + i + "," + j + ",_1列");
                tableHeader11.setField("field" + i + j + "_1[1][0].name");
                tableHeader11.setWidth(20);
                children.add(tableHeader11);

                TableHeader tableHeader1 = new TableHeader();
                tableHeader1.setHeaderText("第" + i + "," + j + "列");
                tableHeader1.setField("field" + i + j);

                List<TableHeader> children2 = new ArrayList<>();
                for (int k = 0; k < 3; k++) {
                    TableHeader tableHeader2 = new TableHeader();
                    tableHeader2.setHeaderText("第" + i + "," + j + "," + k + "列");
                    tableHeader2.setField("field" + i + j + k);
                    tableHeader2.setWidth(50);
                    children2.add(tableHeader2);
                }
                tableHeader1.setChildren(children2);
                children.add(tableHeader1);
            }
            tableHeader.setChildren(children);
            tableHeaderList.add(tableHeader);
        }

        List<Map<String, Object>> tableData = new ArrayList<>();
        Map dataMap = new HashMap();

        Map<String, String> aaa = new HashMap<>();
        aaa.put("name", "name是谁");
        Map<String, Map<String, String>> bbb = new HashMap<>();
        bbb.put("field_1_1_1", aaa);
        List<Map<String, Map<String, String>>> ccc = new ArrayList<>();
        ccc.add(bbb);
        B zzz = new B();
        zzz.setData(ccc);
        dataMap.put("field_1", zzz);

        A hhh = new A();
        hhh.setName("name是你");
        A iii = new A();
        iii.setName("name是我");
        A jjj = new A();
        jjj.setName("name是他");
        List<A> ddd = new ArrayList<>();
        ddd.add(hhh);
        ddd.add(iii);
        ddd.add(jjj);

        A kkk = new A();
        kkk.setName("name是谁");
        A lll = new A();
        lll.setName("name是WHO");
        List<A> eee = new ArrayList<>();
        eee.add(kkk);
        eee.add(lll);
        List<List<A>> fff = new ArrayList<>();
        fff.add(ddd);
        fff.add(eee);
        dataMap.put("field00_1", fff);
        dataMap.put("field000", "field000value");
        dataMap.put("field001", "field001value");
        dataMap.put("field002", "field002value");
        dataMap.put("field01_1", fff);
        dataMap.put("field010", "field010value");
        dataMap.put("field011", "field011value");
        dataMap.put("field012", "field012value");
        dataMap.put("field02_1", fff);
        dataMap.put("field020", "field020value");
        dataMap.put("field021", "field021value");
        dataMap.put("field022", "field022value");
        tableData.add(dataMap);
        Map dataMap1 = new HashMap();
        dataMap1.put("field_1", zzz);
        dataMap1.put("field00_1", fff);
        dataMap1.put("field000", "field000value");
        dataMap1.put("field001", "field001value$bg[#4394ff]");
        dataMap1.put("field002", "field002value");
        dataMap1.put("field01_1", fff);
        dataMap1.put("field010", "field010value$bg[#4394ff]");
        dataMap1.put("field011", "field011value");
        dataMap1.put("field012", "field012value$bg[#4394ff]");
        dataMap1.put("field02_1", fff);
        dataMap1.put("field020", "field020value");
        dataMap1.put("field021", "field021value");
        dataMap1.put("field022", "field022value");
        tableData.add(dataMap1);

        List<Map<String, String>> tableData1 = new ArrayList<>();
        Map dataMap2 = new HashMap();
        dataMap2.put("field_1", zzz);
        dataMap2.put("field00_1", fff);
        dataMap2.put("field000", "field000value");
        dataMap2.put("field001", "field001value");
        dataMap2.put("field002", "field002value");
        dataMap2.put("field01_1", fff);
        dataMap2.put("field010", "field010value");
        dataMap2.put("field011", "field011value");
        dataMap2.put("field012", "field012value");
        dataMap2.put("field02_1", fff);
        dataMap2.put("field020", "field020value");
        dataMap2.put("field021", "field021value");
        dataMap2.put("field022", "field022value");
        tableData1.add(dataMap2);

        Font tableHeaderFont = exportExcel.createTableHeaderFont();
        tableHeaderFont.setFontName("微软雅黑");
        tableHeaderFont.setFontHeightInPoints((short) 20);
        exportExcel.setTableHeaderFont(tableHeaderFont);
        exportExcel.drawTable(tableHeaderList, tableData, 1, 1);
        exportExcel.drawTable(tableHeaderList, tableData1, 1, 20);
        exportExcel.write("D:\\test.xlsx");
    }

    public static void main(String[] args) throws IOException {
        List<TableHeader> tableHeaderList = new ArrayList<>();
        for (int j = 0; j < 20; j++) {
            TableHeader tableHeader = new TableHeader();
            tableHeader.setHeaderText("第" + j + "列");
            tableHeader.setAlign("center");
            tableHeaderList.add(tableHeader);
        }

        List<List<String>> tableData = new ArrayList<>();
        for (int i = 0; i < 54812; i++) {
            List<String> rowData = new ArrayList<>();
            for (int j = 0; j < 20; j++) {
                rowData.add("单元格" + i + "," + j);
            }
            tableData.add(rowData);
        }

        NomalExportExcel nomalExportExcel = new NomalExportExcel(tableHeaderList, tableData);

        nomalExportExcel.export("D:\\test", "test");
    }
}

class A{
    private String name;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}

class B{
    private List<Map<String, Map<String, String>>> data;

    public List<Map<String, Map<String, String>>> getData() {
        return data;
    }

    public void setData(List<Map<String, Map<String, String>>> data) {
        this.data = data;
    }
}