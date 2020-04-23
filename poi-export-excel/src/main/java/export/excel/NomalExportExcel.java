package export.excel;

import export.entity.TableHeader;
import net.lingala.zip4j.ZipFile;
import net.lingala.zip4j.model.ZipParameters;
import net.lingala.zip4j.model.enums.CompressionLevel;
import net.lingala.zip4j.model.enums.CompressionMethod;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @param <T>
 * @author deng-zj
 * @date 2020-04-23
 * @description 普通导出Excel，只需要传入表头、表体。使用默认的样式导出Excel。适用于最基本的导出
 * @since JDK1.8
 */
public class NomalExportExcel<T> {

    private static final Logger logger = LoggerFactory.getLogger(NomalExportExcel.class);

    private List<TableHeader> tableHeaderList;
    private List<T> tableData;

    public NomalExportExcel(List<TableHeader> tableHeaderList, List<T> tableData) {
        this.tableHeaderList = tableHeaderList;
        this.tableData = tableData;
    }

    /**
     * 导出Excel
     *
     * @param fileDir
     * @param fileName
     * @return 返回导出后最终的文件路径
     */
    public String export(String fileDir, String fileName) throws IOException {
        logger.info("==================================开始导出Excel");
        Long exportStartTime = System.currentTimeMillis();

        /**
         * 将导出数据进行分sheet分excel，数据大的情况下将数据分成多个Excel保存数据，一个Excel又分多个Sheet
         */
        List<List<List<T>>> allExcelData = new ArrayList<>();
        List<List<T>> excelData = new ArrayList<>();
        List<T> sheetData = new ArrayList<>();
        for (int i = 0; i < this.tableData.size(); i++) {
            if (i != 0 && i % 100 == 0) {
                excelData.add(sheetData);
                sheetData = new ArrayList<>();

                if (excelData.size() % 10 == 0) {
                    allExcelData.add(excelData);
                    excelData = new ArrayList<>();
                }
            }
            sheetData.add(this.tableData.get(i));

            if (i == this.tableData.size() - 1 && i % 100 != 0) {
                excelData.add(sheetData);
                allExcelData.add(excelData);
            }
        }

        /**
         * 将分好的数据进行导出
         */
        String xlsxFilePath = fileDir + File.separator + fileName + ".xlsx";
        for (int i = 0; i < allExcelData.size(); i++) {
            logger.info("==================================开始导出第" + (i + 1) + "个Excel");
            Long exportExcelStartTime = System.currentTimeMillis();

            ExportExcel exportExcel = new ExportExcel();

            /**
             * 创建Excel
             */
            excelData = allExcelData.get(i);
            XSSFWorkbook workbook = exportExcel.createExcel(excelData.size(), fileName);
            exportExcel.createTableHeaderFont();

            for (int j = 0; j < excelData.size(); j++) {
                XSSFSheet sheet = workbook.getSheetAt(j);
                sheet.setDefaultRowHeight((short) (255 * 2));
                exportExcel.setSheet(sheet);
                exportExcel.setNowMaxRowNums(0);
                exportExcel.setNowMaxColNums(0);
                exportExcel.drawTable(this.tableHeaderList, excelData.get(j));
            }

            File exportDir = new File(fileDir);
            if (!exportDir.exists()) {
                exportDir.mkdirs();
            }

            if (allExcelData.size() > 1) {
                exportExcel.write(fileDir + File.separator + fileName + "(" + (i + 1) + ").xlsx");
            } else {
                exportExcel.write(xlsxFilePath);
            }

            Long exportExcelEndTime = System.currentTimeMillis();
            logger.info("==================================第" + (i + 1) + "个Excel导出完成，共耗时：" + (exportExcelEndTime - exportExcelStartTime) + "ms");
        }

        Long exportEndTime = System.currentTimeMillis();
        logger.info("==================================Excel全部导出成功。共" + allExcelData.size() + "个Excel，耗时" + (exportEndTime - exportStartTime) + "ms");

        if (allExcelData.size() > 1) {
            /**
             * 导出多个文件时，将导出的所有文件进行压缩至一个压缩包
             */
            ZipParameters parameters = new ZipParameters();
            parameters.setCompressionMethod(CompressionMethod.DEFLATE); // 压缩方式
            parameters.setCompressionLevel(CompressionLevel.NORMAL); // 压缩级别

            String zipFilePath = fileDir + File.separator + fileName + ".zip";
            File file = new File(zipFilePath);
            if(file.exists()){
                file.delete();
            }

            ZipFile zipFile = new ZipFile(zipFilePath);

            File exportDir = new File(fileDir);
            File[] exportFiles = exportDir.listFiles();
            for (File exportFile : exportFiles) {
                if (exportFile.getName().startsWith(fileName) && exportFile.getName().endsWith(".xlsx")) {
                    zipFile.addFile(exportFile);

                    /**
                     * 将原Excel文件删除，只留下压缩包
                     */
                    exportFile.delete();
                }
            }

            return zipFilePath;
        }
        return xlsxFilePath;
    }
}
