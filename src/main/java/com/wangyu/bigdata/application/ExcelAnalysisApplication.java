package com.wangyu.bigdata.application;

import com.wangyu.bigdata.application.poi.ReadExcel;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.*;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

/**
 * @description: 解析Excel里相关字段的值形成对应的处理DML语句
 * @author: wangyu
 * @create: 2020-07-15 18:04
 **/
public class ExcelAnalysisApplication {

    private static final Logger logger = LoggerFactory.getLogger(ExcelAnalysisApplication.class);

    private static final String XLS_FILE_SUFFIX = ".xls";

    private static final String XLSX_FILE_SUFFIX = ".xlsx";

    /**
     * @Description: 已经处理数据量
     * @Author: wangyu
     * @Date: 2020/7/1
     */
    private static AtomicInteger processnumber = new AtomicInteger(0);
    /**
     * @Description: 读到全部数据量
     * @Author: wangyu
     * @Date: 2020/7/1
     */
    private static AtomicInteger allnumber = new AtomicInteger(0);


    public static void main(String[] args) {

        if (args[0] == null && args[1] == null) {
            System.out.println("输入/输出文件路径为空，程序退出");
            logger.error("输入/输出文件路径为空，程序退出");
            System.exit(-1);
        }
        // 设置要读取的文件路径
        String filepath = args[0];
        // 设置要储存的文件路径
        String outfilepath = args[1];
        // 处理Excel数据
        excelToJson(filepath, outfilepath);
        System.out.println("程序运行结束");
        logger.info("程序运行结束");

    }

    /**
     * @Description: 解析Excel里边的内容，并生成相关语句语句
     * @Author: wangyu
     * @Date: 2020/7/18
     */
    private static void excelToJson(String filepath, String outfilepath) {
        // 构造一个Scanner对象，其传入参数System.in
        Scanner scan = new Scanner(System.in);
        // 获取工作簿
        Workbook workbook = null;
        List<Map<String, String>> maps = new ArrayList<>();
        try {
            InputStream instream = new FileInputStream(filepath);
            File file = new File(outfilepath);
            //如果没有文件就创建
            if (!file.isFile()) {
                file.createNewFile();
            }
            BufferedWriter writer = new BufferedWriter(new FileWriter(outfilepath));
            // 截取文件后缀名
            String filetring = filepath.substring(filepath.lastIndexOf("."));
            //HSSFWorkbook相当于一个excel文件，HSSFWorkbook是解析excel2007之前的版本（.xls） 之后版本使用XSSFWorkbook（.xlsx）
            if (XLS_FILE_SUFFIX.equals(filetring)) {
                workbook = new HSSFWorkbook(instream);
            } else if (XLSX_FILE_SUFFIX.equals(filetring)) {
                workbook = new XSSFWorkbook(instream);
            }
            System.out.println("程序检查到您的文件为：【" + filetring + "】类型 ");
            logger.info("程序检查到您的文件为：【" + filetring + "】类型 ");
            if (workbook != null) {
                for (int sheetat = 0; sheetat < workbook.getNumberOfSheets(); sheetat++) {
                    // 获得sheet工作簿
                    Sheet sheet = workbook.getSheetAt(sheetat);
                    // 获取本sheet行数
                    addAllNumber(sheet.getLastRowNum());
                    // 获取表头
                    Row tableheader = sheet.getRow(0);
                    // 创建表头数组
                    String[] tableheaders = new String[tableheader.getPhysicalNumberOfCells()];
                    System.out.println("查询到的字段有：");
                    logger.info("查询到的字段有：");
                    // 遍历表头储存到数组里
                    for (int tableheaderindex = 0; tableheaderindex < tableheader.getLastCellNum(); tableheaderindex++) {
                        tableheaders[tableheaderindex] = String.valueOf(tableheader.getCell(tableheaderindex));
                        System.out.println(tableheaderindex + ": " + tableheaders[tableheaderindex]);
                        logger.info(tableheaderindex + ": " + tableheaders[tableheaderindex]);
                    }
                    System.out.println("请输入要处理的字段索引(大于1个请用逗号隔开)");
                    String scanString = scan.next();
                    List<Integer> parseIndex = new ParseScanString().parseIndex(scanString, tableheaders.length);
                    logger.info("输入要处理的字段索引:" + parseIndex.toString());
                    System.out.println("请输入要执行语句作为set的字段索引(大于1个请用逗号隔开)");
                    String setFieldString = scan.next();
                    List<Integer> updateIndex = new ParseScanString().parseIndex(setFieldString, tableheaders.length);
                    logger.info("输入要执行语句作为set的字段索引:" + updateIndex.toString());
                    System.out.println("请输入要执行语句作为where条件的字段索引(大于1个请用逗号隔开)");
                    String whereFieldString = scan.next();
                    List<Integer> whereIndex = new ParseScanString().parseIndex(whereFieldString, tableheaders.length);
                    logger.info("输入要执行语句作为where条件的字段索引:" + whereIndex.toString());
                    System.out.println("请输入要执行语句的表名");
                    String tableString = scan.next();
                    logger.info("输入要执行语句的表名:" + tableString);
                    for (int rownum = 1; rownum < sheet.getPhysicalNumberOfRows(); rownum++) {
                        Row row = sheet.getRow(rownum);
                        //每一行创建一个Map
                        Map<String, String> rowmap = new HashMap<>();
                        for (int colnumindex = 0; colnumindex < row.getPhysicalNumberOfCells(); colnumindex++) {
                            if (parseIndex.contains(colnumindex)) {
                                String cell = ReadExcel.getvalue(sheet, rownum, colnumindex);
                                rowmap.put(tableheaders[colnumindex], cell);
                            }
                        }
                        maps.add(rowmap);
                    }
                    List<StringBuilder> stringBuilders = new MapListToUpdateDml().parseToDml(tableheaders, updateIndex, whereIndex, maps, tableString);
                    // 写入文件
                    for (StringBuilder stringBuilder : stringBuilders) {
                        writer.write(stringBuilder + "\r\n");
                    }
                    writer.close();
                }
            }
            instream.close();
        } catch (Exception e) {
            logger.error("处理Excel数据出现错误，错误信息：【{}】", e.getMessage(), e);
            e.printStackTrace();
        }
    }


    /**
     * @Description: 将Excel解析好的数据传换成DML的接口
     * @Author: wangyu
     * @Date: 2020/7/18
     */
    interface MapListToDml {
        List<StringBuilder> parseToDml(String[] tableheaders, List<Integer> updateIndexs,
                                       List<Integer> whereIndexs, List<Map<String, String>> maps, String tableString);
    }

    /**
     * @Description: 将Excel解析出的MapList数据转成update语句
     * @Param: tableheaders：表头数组 updateIndexs：需要update的字段索引List whereIndexs：需要做where条件的索引List maps：处理的结果集 tableString:表名
     * @return: List<StringBuilder> 结果List
     * @Author: wangyu
     * @Date: 2020/7/17
     */
    static class MapListToUpdateDml implements MapListToDml {

        @Override
        public List<StringBuilder> parseToDml(String[] tableheaders, List<Integer> updateIndexs, List<Integer> whereIndexs, List<Map<String, String>> maps, String tableString) {
            System.out.println("开始执行生成update语句");
            logger.info("开始执行生成update语句");
            List<StringBuilder> updateSentences = new ArrayList<>();

            for (Map<String, String> map : maps) {
                StringBuilder updateSentence = new StringBuilder();
                updateSentence.append("update " + tableString + " set ");
                for (int i = 0; i < updateIndexs.size(); i++) {
                    int updateIndex = updateIndexs.get(i);
                    if (i > 0) {
                        updateSentence.append(",");
                    }
                    String updateValue = map.get(tableheaders[updateIndex]);
                    updateSentence.append(tableheaders[updateIndex] + " = " + updateValue);
                }
                updateSentence.append(" where ");
                for (int j = 0; j < whereIndexs.size(); j++) {
                    int whereIndex = whereIndexs.get(j);
                    if (j > 0) {
                        updateSentence.append(" and ");
                    }
                    String wherevalue = map.get(tableheaders[whereIndex]);
                    updateSentence.append(tableheaders[whereIndex] + " = " + wherevalue);
                }
                updateSentence.append(";");
                System.out.println(updateSentence);
                updateSentences.add(updateSentence);
                addProcessNumber(1);
            }
            return updateSentences;
        }
    }


    /**
     * @Description: 定义处理表头索引字符串的接口
     * @Param: scanString：用户输入的索引字符串  tableheaderslength：表头的长度
     * @return: List<Integer> 索引List
     * @Author: wangyu
     * @Date: 2020/7/17
     */
    interface ParseString {
        List<Integer> parseIndex(String scanString, Integer tableheaderslength);

    }

    /**
     * @Description: 实现解析用户输入的索引字符串成索引List
     * @Param: scanString：用户输入的索引字符串  tableheaderslength：表头的长度
     * @return: List<Integer> 索引List
     * @Author: wangyu
     * @Date: 2020/7/17
     */
    static class ParseScanString implements ParseString {

        @Override
        public List<Integer> parseIndex(String scanString, Integer tableheaderslength) {
            List<Integer> splitcolnum = new ArrayList<>();
            if (scanString.contains(",")) {
                splitcolnum = Arrays.asList(scanString.split(",")).stream().map(s -> Integer.parseInt(s.trim())).filter(p -> p < tableheaderslength).collect(Collectors.toList());
            } else if (Integer.parseInt(scanString) < tableheaderslength) {
                splitcolnum.add(Integer.parseInt(scanString));
            } else {
                System.out.println("您输入的字段索引错误，请重新启动程序");
                logger.info("输入的字段索引错误，停止程序");
                System.exit(-1);
            }
            return splitcolnum;
        }
    }

    /**
     * @Description: 处理总量进行加总操作
     * @Author: wangyu
     * @Date: 2020/7/3
     */
    private synchronized static void addProcessNumber(Integer number) {
        logger.info(Thread.currentThread().getId() + " 处理进度(已处理数/总数):【{}】", processnumber.addAndGet(number) + "/" + allnumber);
    }

    /**
     * @Description: 查询总量进行加总操作
     * @Author: wangyu
     * @Date: 2020/7/3
     */
    private synchronized static void addAllNumber(Integer number) {
        logger.info(Thread.currentThread().getId() + " 查询到的总数:【{}】", allnumber.addAndGet(number));
    }


}
