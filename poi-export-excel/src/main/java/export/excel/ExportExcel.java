package export.excel;

import export.entity.TableHeader;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author deng-zj
 * @date 2020-04-23
 * @description 利用POI进行自定义导出Excel
 *
 * 使用说明：
 * 1、创建ExportExcel对象
 * 2、用ExportExcel实例对象的方法<method>createExcel</method>创建Excel对象
 * 3、画表格之前可以对表格进行设置样式
 * @see <method>createTableHeaderFont</method> 获取表头字体样式对象，对表头内容进行自定义字体样式
 * @see <method>createTableBodyFont</method> 获取表体字体样式对象，对表体内容进行自定义字体样式
 * @see <method>setAddBorder</method> 设置表格是否添加边框
 * @see <method>setAddTableHeaderBorder</method> 设置表格的表头是否添加边框
 * 4、组装表头结构[对于多级表头会进行合并居中，如果是map或者实体类，field字段一定要匹配。否则取不到数据就会抛出异常]
 * @see TableHeader 表头对象 注：field支持多级取值 例如：a.b[n].c | a.b.c | a[n].b.c
 * 5、将表格填充至Excel
 * @see <method>drawTable</method> 这个方法提供了两个实现方式：
 * 第一种：以追加的形式将表格填充至Excel，两个表格之间默认间隔两行。
 * 第二种：以指定下标的形式将表格填充至Excel指定位置。
 * 6、将Excel写出至指定磁盘路径
 *
 * 这里额外提供了一个方法<method>setCellBackGround</method>可以设置指定区域内所有单元格的背景色
 *
 * @since   JDK1.8
 * @param <T>
 */
public class ExportExcel<T> {

    private static final Logger logger = LoggerFactory.getLogger(ExportExcel.class);

    /**
     * 表格间距
     */
    private static final Integer SPACING_BETWEEN_TABLES = 2;

    /**
     * 默认<行>高度
     */
    private static final Short DEFULT_ROW_HEIGHT = 255 * 2;

    private XSSFWorkbook workBook;
    private XSSFSheet sheet;
    private Font tableHeaderFont;
    private Font tableBodyFont;

    private List<TableHeader> tableHeaderList;
    private List<T> tableData;
    private Integer nowMaxRowNums = 0;
    private Integer nowMaxColNums = 0;

    private Integer tableMaxRows = 0;
    private Integer tableMaxCols = 0;

    private Integer tableHeaderRowNum = 0;

    /**
     * 是否添加边框
     */
    private boolean addBorder = true;

    /**
     * 是否添加表头边框
     */
    private boolean addTableHeaderBorder = true;

    private Integer tableCount = 0;

    /**
     * 创建Excel
     *
     * @param sheetName
     * @return
     */
    public XSSFWorkbook createExcel(String sheetName) {
        logger.info("=================开始创建Excel(" + sheetName + ")工作簿=================");
        this.workBook = new XSSFWorkbook();
        this.workBook.createSheet(sheetName);
        this.sheet = this.workBook.getSheet(sheetName);

        this.sheet.setDefaultRowHeight(DEFULT_ROW_HEIGHT);
        logger.info("=================(" + sheetName + ")工作簿创建完成=================");
        return this.workBook;
    }

    /**
     * 传入表格数据，将数据追加至Excel中
     *
     * @param tableHeaderList 表头数据
     * @param tableData       表格数据
     * @return
     */
    public XSSFWorkbook drawTable(List<TableHeader> tableHeaderList, List<T> tableData) {
        logger.info("=================开始绘制第" + (++tableCount) + "个表格=================");
        Long startTime = System.currentTimeMillis();

        this.tableHeaderList = tableHeaderList;
        this.tableData = tableData;
        this.tableHeaderRowNum = 0;

        Integer startRowIndex;
        if (this.nowMaxRowNums == 0) {
            startRowIndex = this.nowMaxRowNums;
        } else {
            startRowIndex = this.nowMaxRowNums + SPACING_BETWEEN_TABLES;
        }

        Integer tableHeaderColNum = this.getTableHeaderColumn(this.tableHeaderList, new ArrayList<>()).size();
        if (tableHeaderColNum > this.nowMaxColNums) {
            this.nowMaxColNums = tableHeaderColNum;
        }

        if (this.nowMaxRowNums < startRowIndex) {
            this.fillCol(startRowIndex);
        }

        /**
         * 开始绘制表头数据
         * 因为不需要传入<列>下标，默认从第一列开始绘制
         */
        this.drawTableHeader(startRowIndex, 0, this.tableHeaderList, 0, 0);

        /**
         * 表格最大列下标 = 表格总字段数
         * 表格最大行下标 = 开始行 + 表头总行数 + 表格数据总行数
         */
        this.tableMaxCols = this.getTableHeaderColumn(this.tableHeaderList, new ArrayList<>()).size();
        this.tableMaxRows = startRowIndex + this.tableHeaderRowNum + 1 + this.tableData.size();

        /**
         * 多级表头需要填充<列>，然后进行合并
         */
        this.fillCol(this.nowMaxRowNums);
        this.mergeCol(startRowIndex, 0);
        this.mergeRow(startRowIndex, 0);

        /**
         * 开始绘制表格数据
         */
        this.drawTableData(startRowIndex, 0);

        this.nowMaxRowNums = this.sheet.getPhysicalNumberOfRows();
        if (this.nowMaxRowNums != 0) {
            this.nowMaxColNums = this.sheet.getRow(0).getPhysicalNumberOfCells();
        }

        /**
         * 添加边框
         */
        if (addBorder) {
            this.addBorder(startRowIndex, 0);
        }

        Long endTime = System.currentTimeMillis();
        logger.info("=================第" + tableCount + "个表格绘制完成。耗时" + (endTime - startTime) + "ms=================");
        return this.workBook;
    }

    /**
     * 传入Row开始坐标和Col开始坐标进行绘制表格
     *
     * @param tableHeaderList 表头数据
     * @param tableData       表格数据
     * @param startRowIndex   <行>开始坐标
     * @param startColIndex   <列>开始坐标
     * @return
     */
    public XSSFWorkbook drawTable(List<TableHeader> tableHeaderList, List<T> tableData, Integer startRowIndex, Integer startColIndex) {
        logger.info("=================开始绘制第" + (++tableCount) + "个表格=================");
        Long startTime = System.currentTimeMillis();

        this.tableHeaderList = tableHeaderList;
        this.tableData = tableData;
        this.tableHeaderRowNum = 0;

        Integer tableHeaderColNum = this.getTableHeaderColumn(this.tableHeaderList, new ArrayList<>()).size();
        if ((tableHeaderColNum + startColIndex) > this.nowMaxColNums) {
            this.nowMaxColNums = tableHeaderColNum + startColIndex;
        }

        if (this.nowMaxRowNums < startRowIndex) {
            this.fillCol(startRowIndex);
        }

        /**
         * 开始绘制表头数据
         */
        this.drawTableHeader(startRowIndex, startColIndex, this.tableHeaderList, 0, 0);

        /**
         * 表格最大列下标 = 开始列 + 表格总字段数
         * 表格最大行下标 = 开始行 + 表头总行数 + 表格数据总行数
         */
        this.tableMaxCols = startColIndex + this.getTableHeaderColumn(this.tableHeaderList, new ArrayList<>()).size();
        this.tableMaxRows = startRowIndex + this.tableHeaderRowNum + 1 + this.tableData.size();

        /**
         * 多级表头需要填充<列>，然后进行合并
         * 开始行下标+表头行数（this.tableHeaderRowNum是下标）要得到实际数得加1
         */
        this.fillCol(startRowIndex + this.tableHeaderRowNum + 1);
        this.mergeCol(startRowIndex, startColIndex);
        this.mergeRow(startRowIndex, startColIndex);

        /**
         * 开始绘制表格数据
         */
        this.drawTableData(startRowIndex, startColIndex);

        this.nowMaxRowNums = this.sheet.getPhysicalNumberOfRows();
        if (this.nowMaxRowNums != 0) {
            this.nowMaxColNums = this.sheet.getRow(0).getPhysicalNumberOfCells();
        }

        /**
         * 表格绘制完成后，填充Excel表格缺少的单元格
         */
        this.fillCol(this.nowMaxRowNums);

        /**
         * 添加边框
         */
        if (addBorder) {
            this.addBorder(startRowIndex, startColIndex);
        }

        Long endTime = System.currentTimeMillis();
        logger.info("=================第" + tableCount + "个表格绘制完成。耗时" + (endTime - startTime) + "ms=================");
        return this.workBook;
    }

    /**
     * 绘制表头
     *
     * @param rowIndex        开始<行>下标
     * @param colIndex        开始<列>下标
     * @param tableHeaderList 表头数组
     * @param rowCount        行计数
     * @param colCount        列计数
     */
    private Integer drawTableHeader(Integer rowIndex, Integer colIndex, List<TableHeader> tableHeaderList, Integer rowCount, Integer colCount) {
        for (int i = 0; i < tableHeaderList.size(); i++) {
            XSSFRow row = this.sheet.getRow(rowIndex + rowCount);
            if (row == null) {
                row = this.sheet.createRow(rowIndex + rowCount);
            }
            XSSFCell cell = row.getCell(colIndex + colCount);
            if (cell == null) {
                cell = row.createCell(colIndex + colCount);
            }
            TableHeader tableHeader = tableHeaderList.get(i);
            cell.setCellValue(tableHeader.getHeaderText());

            /**
             * 添加表头字体样式
             */
            XSSFCellStyle cellStyle = cell.getCellStyle().copy();
            cellStyle.setFont(this.tableHeaderFont);

            /**
             * 设置自定义列宽
             */
            this.sheet.setColumnWidth(colIndex + colCount, tableHeader.getWidth() * 255);

            /**
             * 设置自定义背景颜色
             */
            if (!"".equals(tableHeader.getBackground())) {
                this.setCustomBackGround(cellStyle, tableHeader.getBackground());
            }

            /**
             * 设置单元格内容靠左、靠右、居中
             */
            this.setCellAlign(cellStyle, tableHeader);

            cell.setCellStyle(cellStyle);

            if (tableHeader.getChildren() != null) {
                rowCount++;
                if (this.tableHeaderRowNum < rowCount) {
                    this.tableHeaderRowNum = rowCount;
                }
                colCount = drawTableHeader(rowIndex, colIndex, tableHeader.getChildren(), rowCount, colCount);
                rowCount--;
            } else {
                colCount++;
            }
        }

        this.nowMaxRowNums = this.sheet.getPhysicalNumberOfRows();
        return colCount;
    }

    /**
     * 开始绘制表格数据
     * 从表头下绘制表格
     */
    private void drawTableData(Integer startRowIndex, Integer startColIndex) {
        logger.info("=================开始绘制第" + (tableCount) + "个表格中的数据=================");

        List<TableHeader> tableHeaderColumnList = getTableHeaderColumn(this.tableHeaderList, new ArrayList<>());
        for (int i = 0; i < this.tableData.size(); i++) {
            T rowData = (T) this.tableData.get(i);

            /**
             * 创建或者获取行对象
             * 开始行下标+表头行数（this.tableHeaderRowNum是下标）要得到实际数得加1
             */
            XSSFRow row = this.sheet.getRow(startRowIndex + this.tableHeaderRowNum + 1 + i);
            if (row == null) {
                row = this.sheet.createRow(startRowIndex + this.tableHeaderRowNum + 1 + i);
            }

            for (int j = 0; j < tableHeaderColumnList.size(); j++) {

                /**
                 * 创建或者获取列对象
                 */
                XSSFCell cell = row.getCell(startColIndex + j);
                if (cell == null) {
                    cell = row.createCell(startColIndex + j);
                }

                /**
                 * 设置自定义文字样式
                 */
                XSSFCellStyle cellStyle = cell.getCellStyle().copy();
                cellStyle.setFont(this.tableBodyFont);

                /**
                 * 将数据放入Excel单元格
                 */
                try {
                    String[] contentAndColor = new String[]{""}; // 内容和颜色
                    String field = tableHeaderColumnList.get(j).getField();

                    String[] arrMultistageField = field.split("\\.");
                    /**
                     * 如果数据取值字段是一级的情况下，则直接取值
                     * 否则，则一级一级往下取值
                     */
                    if (arrMultistageField.length == 1) {
                        if (rowData instanceof Map) {
                            contentAndColor = ((Map) rowData).get(field).toString().split("\\$bg");
                        } else if (rowData instanceof List) {
                            contentAndColor = ((List) rowData).get(j).toString().split("\\$bg");
                        } else {
                            Method[] methods = rowData.getClass().getMethods();
                            for (Method method : methods) {
                                if (("get" + field).toLowerCase().equals(method.getName().toLowerCase())) {
                                    contentAndColor = method.invoke(rowData).toString().split("\\$bg");
                                    break;
                                }
                            }
                        }
                    } else if (arrMultistageField.length > 1) {
                        /**
                         * 获取多级字段的数据
                         */
                        contentAndColor = this.getMultistageFieldData(rowData, contentAndColor, arrMultistageField);
                    }
                    cell.setCellValue(contentAndColor[0]);

                    /**
                     * 设置自定义背景颜色
                     */
                    if (contentAndColor.length > 1) {
                        this.setCustomBackGround(cellStyle, contentAndColor[1]);
                    }

                    /**
                     * 设置单元格内容靠左、靠右、居中、是否换行
                     */
                    this.setCellAlign(cellStyle, tableHeaderColumnList.get(j));

                    cell.setCellStyle(cellStyle);
                } catch (Exception e) {
                    logger.error("导出数据格式异常，请确认TableHeader中field与导出数据的field一致！\t" + e.getLocalizedMessage());
                    e.printStackTrace();
                }
            }
        }

        logger.info("=================第" + (tableCount) + "个表格中的数据绘制完成=================");
    }

    /**
     * 获取多级字段的数据
     *
     * @param rowData            表格一行数据
     * @param contentAndColor    数据存放位置
     * @param arrMultistageField 多级字段
     * @return
     */
    private String[] getMultistageFieldData(T rowData, String[] contentAndColor, String[] arrMultistageField) throws InvocationTargetException, IllegalAccessException {
        /**
         * 定义不确定 data 类型, data可能是Map或者List
         * 定义 k 变量, 如果第一个key值对应的是List，k就从零开始，原因是List不占取值变量名，用a[n]取值
         * 如果第一个key值是Map， k就从1开始，原因是Map会占用取值变量名, 用a.b取值
         * 定义 firstKey 变量，如果第一个key值对应的是List，取数组第一个下标的前缀变量作为key
         * 如果第一个key值是Map，取值数组第一个下标作为key
         */
        T data;
        int k = 0;
        String firstKey;
        if (arrMultistageField[0].endsWith("]")) {
            firstKey = arrMultistageField[0].split("\\[")[0];
        } else {
            firstKey = arrMultistageField[0];
            k = 1;
        }

        data = (T) ((Map) rowData).get(firstKey);
        /**
         * 如果通过Map没有取到数据，则当做对象通过get方法进行取值
         */
        if (data == null) {
            Method[] methods = data.getClass().getMethods();
            for (Method method : methods) {
                if (("get" + firstKey).toLowerCase().equals(method.getName().toLowerCase())) {
                    data = (T) method.invoke(data);
                    break;
                }
            }
        }

        for (; k < arrMultistageField.length; k++) {
            if (arrMultistageField[k].endsWith("]")) {
                String[] fieldAndIndex = arrMultistageField[k].split("\\[");

                /**
                 * 如果是Map，直接用field进行取值
                 * 如果不是Map也不是List，则当做对象通过get方法进行取值
                 */
                if (data instanceof Map) {
                    data = (T) ((Map) data).get(arrMultistageField[k].split("\\[")[0]);
                } else if (!(data instanceof List)) {
                    Method[] methods = data.getClass().getMethods();
                    for (Method method : methods) {
                        if (("get" + arrMultistageField[k].split("\\[")[0]).toLowerCase().equals(method.getName().toLowerCase())) {
                            data = (T) method.invoke(data);
                            break;
                        }
                    }
                }

                /**
                 * 将取出的List通过field传入下标进行取值
                 */
                for (int l = 1; l < fieldAndIndex.length; l++) {
                    data = (T) ((List) data).get(Integer.parseInt(fieldAndIndex[l].substring(0, fieldAndIndex[l].length() - 1)));
                }
            } else {
                /**
                 * 如果是Map，直接用field进行取值
                 * 如果不是Map，则当做对象通过get方法进行取值
                 */
                if (data instanceof Map) {
                    data = (T) ((Map) data).get(arrMultistageField[k]);
                } else {
                    Method[] methods = data.getClass().getMethods();
                    for (Method method : methods) {
                        if (("get" + arrMultistageField[k]).toLowerCase().equals(method.getName().toLowerCase())) {
                            data = (T) method.invoke(data);
                            break;
                        }
                    }
                }
            }
        }

        /**
         * 数据取完之后，字符串数据放入内容和颜色的数组中
         */
        if (data instanceof String) {
            contentAndColor = ((String) data).split("\\$bg");
        }
        return contentAndColor;
    }

    /**
     * 填充列
     *
     * @param maxRows
     */
    private void fillCol(Integer maxRows) {
        /**
         * 填充列，判断当前<行>中是否缺少<列>，缺少就创建。
         */
        for (int i = 0; i < maxRows; i++) {
            XSSFRow row = this.sheet.getRow(i);
            if (row == null) {
                row = this.sheet.createRow(i);
            }

            for (int j = 0; j < this.nowMaxColNums; j++) {
                XSSFCell cell = row.getCell(j);
                if (cell == null) {
                    row.createCell(j);
                }
            }
        }

        this.nowMaxRowNums = this.sheet.getPhysicalNumberOfRows();
        if (this.nowMaxRowNums != 0) {
            this.nowMaxColNums = this.sheet.getRow(0).getPhysicalNumberOfCells();
        }
    }

    /**
     * 合并列
     *
     * @param startRowIndex
     * @param startColIndex
     */
    private void mergeCol(Integer startRowIndex, Integer startColIndex) {
        int firstRow = startRowIndex, lastRow = startRowIndex, firstCol = startColIndex, lastCol = startColIndex;
        /**
         * 开始行下标+表头行数（this.tableHeaderRowNum是下标）要得到实际数得加1
         */
        for (int i = startRowIndex; i < (startRowIndex + this.tableHeaderRowNum + 1); i++) {
            XSSFRow row = this.sheet.getRow(i);

            firstRow = i;
            lastRow = i;
            firstCol = startColIndex;
            lastCol = startColIndex;
            for (int j = startColIndex; j < this.tableMaxCols; j++) {
                XSSFCell nextCell = row.getCell(j + 1);
                if (nextCell != null) {
                    String nextCellValue = nextCell.getStringCellValue();
                    if ("".equals(nextCellValue)) {
                        /**
                         * 需要判断下一列之前是否有数据，之后的数据不需要判断，第一行不需要判断，如果有数据则不合并。
                         * 这一块的判断主要是给行合并留出空间，不留出空间就是导致需要行合并的数据先被列合并了
                         */
                        boolean isExist = false;
                        for (int k = startRowIndex; k < i; k++) {
                            XSSFRow oldRow = this.sheet.getRow(k);
                            XSSFCell oldCell = oldRow.getCell(j + 1);
                            String oldCellValue = oldCell.getStringCellValue();
                            if (!"".equals(oldCellValue)) {
                                isExist = true;
                                break;
                            }
                        }

                        if (!isExist) {
                            lastCol = j + 1;
                        }
                    } else {
                        /**
                         * 结束列不等于表格开始列下标 并且 开始列等于表格开始列下标 并且 当前行开始列有数据时进行合并并居中（有可能会出现第一列没数据）
                         * 或者
                         * 开始列不等于表格开始列下标 并且 开始列小于结束列时进行合并并居中
                         */
                        if ((lastCol != startColIndex && firstCol == startColIndex && !"".equals(row.getCell(startColIndex).getStringCellValue()))
                                || (firstCol != startColIndex && firstCol < lastCol)) {
                            mergedCenter(firstRow, lastRow, firstCol, lastCol);
                        }
                        firstCol = j + 1;
                    }
                } else {
                    /**
                     * 下一列为null，则是循环结尾，结尾如果存在需要合并的，在额外执行一次合并
                     */
                    if ((lastCol != startColIndex && firstCol == startColIndex && !"".equals(row.getCell(startColIndex).getStringCellValue()))
                            || (firstCol != startColIndex && firstCol < lastCol)) {
                        mergedCenter(firstRow, lastRow, firstCol, lastCol);
                    }
                }
            }
        }
    }

    /**
     * 合并行
     *
     * @param startRowIndex
     * @param startColIndex
     */
    private void mergeRow(Integer startRowIndex, Integer startColIndex) {
        int firstRow = startRowIndex, lastRow = startRowIndex, firstCol = startColIndex, lastCol = startColIndex;
        for (int i = startColIndex; i < this.tableMaxCols; i++) {

            firstRow = startRowIndex;
            lastRow = startRowIndex;
            firstCol = i;
            lastCol = i;
            /**
             * 开始行下标+表头行数（this.tableHeaderRowNum是下标）要得到实际数得加1
             */
            for (int j = startRowIndex; j < (startRowIndex + this.tableHeaderRowNum + 1); j++) {
                XSSFRow nextRow = this.sheet.getRow(j + 1);
                /**
                 * 下一行下标超过本次表格使用的行数时，就不合并了
                 */
                if (nextRow != null && (j + 1) <= (startRowIndex + this.tableHeaderRowNum)) {
                    XSSFCell nextCell = nextRow.getCell(i);
                    if ("".equals(nextCell.getStringCellValue())) {
                        lastRow = j + 1;
                    } else {
                        /**
                         * 结束行不等于表格开始行下标 并且 开始行等于表格开始行下标 并且 当前列开始行有数据时进行合并并居中（有可能会出现第一行没数据）
                         * 或者
                         * 开始行不等于表格开始行下标 并且 开始行小于结束行时进行合并并居中
                         */
                        if ((lastRow != startRowIndex && firstRow == startRowIndex && !"".equals(this.sheet.getRow(startRowIndex).getCell(i).getStringCellValue()))
                                || (firstRow != startRowIndex && firstRow < lastRow)) {
                            mergedCenter(firstRow, lastRow, firstCol, lastCol);
                        }
                        firstRow = j + 1;
                    }
                } else {
                    /**
                     * 下一行为null，则是循环结尾，结尾如果存在需要合并的，在额外执行一次合并
                     */
                    if ((lastRow != startRowIndex && firstRow == startRowIndex && !"".equals(this.sheet.getRow(startRowIndex).getCell(i).getStringCellValue()))
                            || (firstRow != startRowIndex && firstRow < lastRow)) {
                        mergedCenter(firstRow, lastRow, firstCol, lastCol);
                    }
                }
            }
        }
    }

    /**
     * 合并居中
     *
     * @param firstRow 开始行
     * @param lastRow  结束行
     * @param firstCol 开始列
     * @param lastCol  结束行
     */
    private void mergedCenter(int firstRow, int lastRow, int firstCol, int lastCol) {
        // 合并
        this.sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));

        // 居中
        XSSFRow row = this.sheet.getRow(firstRow);
        XSSFCell cell = row.getCell(firstCol);
        XSSFCellStyle cellStyle = cell.getCellStyle().copy();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 设置单元格内容靠左、靠右、居中、换行
     *
     * @param cellStyle
     */
    private void setCellAlign(XSSFCellStyle cellStyle, TableHeader tableHeader) {
        String align = tableHeader.getAlign();
        if ("left".equals(align.toLowerCase())) {
            cellStyle.setAlignment(HorizontalAlignment.LEFT);
        } else if ("right".equals(align.toLowerCase())) {
            cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        } else if ("center".equals(align.toLowerCase())) {
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
        }

        /**
         * 垂直居中
         */
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        /**
         * 文字换行
         */
        cellStyle.setWrapText(tableHeader.getWrapText());

    }

    /**
     * 添加边框
     *
     * @param startRowIndex
     * @param startColIndex
     */
    private void addBorder(Integer startRowIndex, Integer startColIndex) {

        /**
         * 是否添加表头边框
         */
        if (!this.addTableHeaderBorder) {
            startRowIndex += this.tableHeaderRowNum + 1;
        }

        /**
         * 表格全面积设置边框
         * 开始行下标+表头行数（this.tableHeaderRowNum是下标）要得到实际数得加1 + 数据总行数
         */
        for (int i = startRowIndex; i < this.tableMaxRows; i++) {
            for (int j = startColIndex; j < this.tableMaxCols; j++) {
                XSSFCell cell = this.sheet.getRow(i).getCell(j);
                XSSFCellStyle cellStyle = cell.getCellStyle().copy();
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cell.setCellStyle(cellStyle);
            }
        }

    }

    /**
     * 设置指定区域的单元格背景色
     *
     * @param startRowIndex 开始<行>下标
     * @param startColIndex 开始<列>下标
     * @param endRowIndex   结束<行>下标
     * @param endColIndex   结束<列>下标
     */
    public void setCellBackGround(Integer startRowIndex, Integer startColIndex, Integer endRowIndex, Integer endColIndex) {
        endRowIndex = endRowIndex == 0 ? this.nowMaxRowNums - 1 : endRowIndex;
        endColIndex = endColIndex == 0 ? this.nowMaxColNums - 1 : endColIndex;

        for (int i = startRowIndex; i <= endRowIndex; i++) {
            XSSFRow row = this.sheet.getRow(i);
            for (int j = startColIndex; j <= endColIndex; j++) {
                XSSFCell cell = row.getCell(j);

                XSSFCellStyle cellStyle = cell.getCellStyle().copy();
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
                cell.setCellStyle(cellStyle);
            }
        }
    }

    /**
     * 设置自定义背景色
     *
     * @param cellStyle
     * @param strColor
     */
    private void setCustomBackGround(XSSFCellStyle cellStyle, String strColor) {
        if (strColor.lastIndexOf("[") != -1) {
            strColor = strColor.substring(strColor.lastIndexOf("[") + 1, strColor.lastIndexOf("]"));
        }
        String[] strColorRGB = strColor.split(",");

        /**
         * 如果本身就是RGB则不需要转，直接将字符串转成数字类型
         * 如果以#号开始就将十六进制转成RBG
         */
        int[] intColorRGB = null;
        if (strColorRGB.length == 3) {
            intColorRGB = new int[strColorRGB.length];
            for (int k = 0; k < strColorRGB.length; k++) {
                intColorRGB[k] = Integer.parseInt(strColorRGB[k]);
            }
        } else if (strColor.startsWith("#")) {
            intColorRGB = hexToRGB(strColor);
        }

        if (intColorRGB != null) {
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(intColorRGB[0], intColorRGB[1], intColorRGB[2]), new DefaultIndexedColorMap()));
        }
    }

    /**
     * 十六进制转RGB
     *
     * @param hex
     * @return
     */
    private static int[] hexToRGB(String hex) {
        String colorStr = hex;
        if (hex.startsWith("#")) {
            colorStr = hex.substring(1);
        }
        if (StringUtils.length(colorStr) == 8) {
            colorStr = hex.substring(2);
        }

        int r = Integer.valueOf(colorStr.substring(0, 2), 16);
        int g = Integer.valueOf(colorStr.substring(2, 4), 16);
        int b = Integer.valueOf(colorStr.substring(4, 6), 16);
        return new int[]{r, g, b};
    }

    /**
     * 获取表头取值字段
     *
     * @param tableHeaderList       表头集合
     * @param tableHeaderColumnList 表头需要取值的集合
     * @return
     */
    private List<TableHeader> getTableHeaderColumn(List<TableHeader> tableHeaderList, List<TableHeader> tableHeaderColumnList) {
        for (int i = 0; i < tableHeaderList.size(); i++) {
            if (tableHeaderList.get(i).getChildren() != null) {
                tableHeaderColumnList = getTableHeaderColumn(tableHeaderList.get(i).getChildren(), tableHeaderColumnList);
            } else {
                tableHeaderColumnList.add(tableHeaderList.get(i));
            }
        }
        return tableHeaderColumnList;
    }

    /**
     * 创建表头字体样式
     *
     * @return
     */
    public Font createTableHeaderFont() {
        this.tableHeaderFont = this.workBook.createFont();
        this.tableHeaderFont.setBold(true);
        return this.tableHeaderFont;
    }

    /**
     * 创建表格数据字体样式
     *
     * @return
     */
    public Font createTableBodyFont() {
        this.tableBodyFont = this.workBook.createFont();
        return this.tableBodyFont;
    }

    /**
     * 生成并写入Excel
     *
     * @param filePath excel 文件路径（全路径）
     * @throws IOException
     */
    public void write(String filePath) throws IOException {
        this.workBook.write(new FileOutputStream(filePath));
    }

    /*********************************************** get and set method***************************************************/

    public void setTableHeaderFont(Font tableHeaderFont) {
        this.tableHeaderFont = tableHeaderFont;
    }

    public void setTableBodyFont(Font tableBodyFont) {
        this.tableBodyFont = tableBodyFont;
    }

    public boolean isAddBorder() {
        return addBorder;
    }

    public void setAddBorder(boolean addBorder) {
        this.addBorder = addBorder;
    }

    public boolean isAddTableHeaderBorder() {
        return addTableHeaderBorder;
    }

    public void setAddTableHeaderBorder(boolean addTableHeaderBorder) {
        this.addTableHeaderBorder = addTableHeaderBorder;
    }

    public Integer getNowMaxRowNums() {
        return nowMaxRowNums;
    }

    public void setNowMaxRowNums(Integer nowMaxRowNums) {
        this.nowMaxRowNums = nowMaxRowNums;
    }

    public Integer getNowMaxColNums() {
        return nowMaxColNums;
    }

    public void setNowMaxColNums(Integer nowMaxColNums) {
        this.nowMaxColNums = nowMaxColNums;
    }
}
