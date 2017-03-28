package ere.ere.dirful.util;

import java.io.FileInputStream;  
import java.io.FileNotFoundException;  
import java.io.FileOutputStream;  
import java.io.IOException;  
import java.text.DecimalFormat;
import java.util.ArrayList;  
import java.util.Date;
import java.util.HashMap;
import java.util.List;  
  
import org.apache.poi.hssf.usermodel.HSSFCell;  
import org.apache.poi.hssf.usermodel.HSSFCellStyle;  
import org.apache.poi.hssf.usermodel.HSSFFont;  
import org.apache.poi.hssf.usermodel.HSSFRow;  
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
  
  
/** 
 * Java操作Excel封装 
 *  
 * @author Ken 
 * @blog http://blog.csdn.net/arjick/article/details/8182484 
 *  
 */  
public class ExcelUtils {  
  
    /** 
     * 建立excelFile 
     *  
     * @param excelPath 
     * @return 
     */  
    public static boolean createExcelFile(String excelPath) {  
        HSSFWorkbook workbook = new HSSFWorkbook();  
        return outputHSSFWorkbook(workbook, excelPath);  
    }  
  
    /** 
     * 插入新的工作表 
     *  
     * @param excelPath 
     * @param sheetName 
     * @return 
     */  
    public static boolean insertSheet(String excelPath, String sheetName) {  
        try {  
            if (FileUtil.checkFile(excelPath)) {  
                HSSFWorkbook workbook = ExcelUtils.getHSSFWorkbook(excelPath);  
                HSSFSheet sheet = workbook.createSheet(sheetName);  
                return outputHSSFWorkbook(workbook, excelPath);  
            } else {  
                System.out.println(excelPath + ":文件不存在.....");  
            }  
            return false;  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return false;  
    }  
  
    /** 
     * 复制工作表 
     *  
     * @param excelPath 
     * @param sheetName 
     * @param sheetNum 
     * @return 
     */  
    public static boolean copySheet(String excelPath, String sheetName,  
            int formSheetNum) {  
        try {  
            if (FileUtil.checkFile(excelPath)) {  
                HSSFWorkbook workbook = ExcelUtils.getHSSFWorkbook(excelPath);  
                if (!ExcelUtils.checkSheet(workbook, sheetName)) {  
                    workbook.cloneSheet(formSheetNum);  
                    workbook.setSheetName(workbook.getNumberOfSheets() - 1,  
                            sheetName);  
                    return outputHSSFWorkbook(workbook, excelPath);  
                } else {  
                    System.out.println(excelPath + ":存在同名工作表" + sheetName  
                            + ".....");  
                }  
            } else {  
                System.out.println(excelPath + ":文件不存在.....");  
            }  
            return false;  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return false;  
    }  
  
    /** 
     * 插入或者更新单元格 
     *  
     * @param excelPath 
     * @param sheetName 
     * @param rowNum 
     * @param cellNum 
     * @param value 
     * @return 
     */  
    public static boolean insertOrUpdateCell(String excelPath,  
            String sheetName, int rowNum, int cellNum, String value) {  
        try {  
            if (FileUtil.checkFile(excelPath)) {  
                HSSFWorkbook wb = getHSSFWorkbook(excelPath);  
                if (ExcelUtils.checkSheet(wb, sheetName)) {  
                    HSSFSheet sheet = wb.getSheet(sheetName);  
                    HSSFRow row = null;  
                    if (sheet.getLastRowNum() < rowNum) {  
                        row = sheet.createRow(rowNum);  
                    } else {  
                        row = sheet.getRow(rowNum);  
                    }  
                    HSSFCell cell = row.getCell(cellNum);  
                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);  
                    cell.setCellValue(value);  
  
                    return outputHSSFWorkbook(wb, excelPath);  
                } else {  
                    System.out.println(excelPath + "的" + sheetName  
                            + ":工作表不存在.....");  
                }  
            } else {  
                System.out.println(excelPath + ":文件不存在.....");  
            }  
            return false;  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return false;  
    }  
  
    /** 
     * 插入或者更新单元格 
     *  
     * @param excelPath 
     * @param sheetIndex 
     * @param rowNum 
     * @param cellNum 
     * @param value 
     * @return 
     */  
    public static boolean insertOrUpdateCell(String excelPath, int sheetIndex,  
            int rowNum, int cellNum, String value) {  
        try {  
            if (FileUtil.checkFile(excelPath)) {  
                HSSFWorkbook wb = getHSSFWorkbook(excelPath);  
                if (ExcelUtils.checkSheet(wb, sheetIndex)) {  
                    HSSFSheet sheet = wb.getSheetAt(sheetIndex);  
                    HSSFRow row = null;  
                    if (sheet.getLastRowNum() < rowNum) {  
                        row = sheet.createRow(rowNum);  
                    } else {  
                        row = sheet.getRow(rowNum);  
                    }  
                    HSSFCell cell = row.getCell(cellNum);  
                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);  
                    cell.setCellValue(value);  
  
                    return outputHSSFWorkbook(wb, excelPath);  
                } else {  
                    System.out.println(excelPath + "的" + sheetIndex  
                            + ":工作表不存在.....");  
                }  
            } else {  
                System.out.println(excelPath + ":文件不存在.....");  
            }  
            return false;  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return false;  
    }  
  
    /** 
     *  
     * @param excelPath 
     * @param sheetIndex 
     * @param rowNum 
     * @param cellNum 
     * @param value 
     * @param style 
     * @return 
     */  
    public static boolean insertOrUpdateCell(String excelPath, int sheetIndex,  
            int rowNum, int cellNum, String value, HSSFCellStyle style) {  
        try {  
            if (FileUtil.checkFile(excelPath)) {  
                HSSFWorkbook wb = getHSSFWorkbook(excelPath);  
                if (ExcelUtils.checkSheet(wb, sheetIndex)) {  
                    HSSFSheet sheet = wb.getSheetAt(sheetIndex);  
                    HSSFRow row = null;  
                    if (sheet.getLastRowNum() < rowNum) {  
                        row = sheet.createRow(rowNum);  
                    } else {  
                        row = sheet.getRow(rowNum);  
                    }  
                    HSSFCell cell = row.createCell(cellNum);  
                    // cell.setCellType(HSSFCell.CELL_TYPE_STRING);  
  
                    HSSFFont font = wb.createFont();  
                    font.setFontHeight((short) 18);  
  
                    HSSFCellStyle style1 = wb.createCellStyle();  
                    style1.setFont(font);  
                    // cell.setCellStyle(style1);  
                    cell.setCellValue(value);  
                      
                    System.out.println(value);  
                    return outputHSSFWorkbook(wb, excelPath);  
                } else {  
                    System.out.println("insertOrUpdateCell:" + excelPath + "的"  
                            + sheetIndex + ":工作表不存在.....");  
                }  
            } else {  
                System.out.println(excelPath + ":文件不存在.....");  
            }  
            return false;  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return false;  
    }  
  
    /** 
     * 插入指定行数据 
     *  
     * @param excelPath 
     * @param sheetName 
     * @param rowNum 
     * @param values 
     * @return 
     */  
    public static boolean insertOrUpadateRowDatas(String excelPath,  
            String sheetName, int rowNum, String... values) {  
        try {  
            for (int i = 0; i < values.length; i++) {  
                insertOrUpdateCell(excelPath, sheetName, rowNum, i, values[i]);  
            }  
            return false;  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return false;  
    }  
  
    /** 
     * 读入excel 
     *  
     * @param excelPath 
     * @return 
     * @throws Exception 
     */  
    public static HSSFWorkbook getHSSFWorkbook(String excelPath)  
            throws Exception {  
        FileInputStream fs = null;  
        try {  
            fs = new FileInputStream(excelPath);  
            HSSFWorkbook wb = new HSSFWorkbook(fs);  
            return wb;  
        } catch (Exception e) {  
            throw e;  
        } finally {  
            try {  
                if (fs != null)  
                    fs.close();  
            } catch (Exception e) {  
                e.printStackTrace();  
            }  
        }  
    }  
  
    /** 
     * 输入Excel 
     *  
     * @param wb 
     * @param excelPath 
     * @return 
     */  
    private static boolean outputHSSFWorkbook(HSSFWorkbook wb, String excelPath) {  
        FileOutputStream fOut = null;  
        try {  
            fOut = new FileOutputStream(excelPath);  
            wb.write(fOut);  
            fOut.flush();  
            System.out.println(excelPath + ":文件生成...");  
            return true;  
        } catch (FileNotFoundException e) {  
            // TODO 自动生成 catch 块  
            e.printStackTrace();  
        } catch (IOException e) {  
            // TODO 自动生成 catch 块  
            e.printStackTrace();  
        } finally {  
            try {  
                if (fOut != null)  
                    fOut.close();  
            } catch (Exception e) {  
                e.printStackTrace();  
            }  
        }  
        return false;  
  
    }  
  
    /** 
     * 检查是否存在工作表 
     *  
     * @param excelPath 
     * @param sheetName 
     * @return 
     */  
    public static boolean checkSheet(HSSFWorkbook wb, String sheetName) {  
        try {  
            for (int numSheets = 0; numSheets < wb.getNumberOfSheets(); numSheets++) {  
                HSSFSheet sheet = wb.getSheetAt(numSheets);  
                if (sheetName.equals(sheet.getSheetName())) {  
                    return true;  
                }  
            }  
            return false;  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return false;  
    }  
  
    /** 
     * 检查是否存在表数量 
     *  
     * @param wb 
     * @param sheetIndex 
     * @return 
     */  
    public static boolean checkSheet(HSSFWorkbook wb, int sheetIndex) {  
        try {  
            if (wb.getNumberOfSheets() > sheetIndex)  
                return true;  
            return false;  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return false;  
    }  
  
    /** 
     * 清除指定excel的工作表所有内容 
     *  
     * @param excelPath 
     * @param sheet 
     * @return 
     */  
    public static boolean cleanExcelFile1(String excelPath, String sheetName) {  
        try {  
            if (FileUtil.checkFile(excelPath)) {  
                HSSFWorkbook wb = getHSSFWorkbook(excelPath);  
                if (ExcelUtils.checkSheet(wb, sheetName)) {  
                    HSSFSheet sheet = wb.getSheet(sheetName);  
                    for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {  
                        if (sheet.getRow(i) != null)  
                            sheet.removeRow(sheet.getRow(i));  
                    }  
                    return outputHSSFWorkbook(wb, excelPath);  
                } else {  
                    System.out.println("cleanExcelFile:" + excelPath + "的"  
                            + sheetName + ":工作表不存在.....");  
                }  
  
            } else {  
                System.out.println(excelPath + ":文件不存在.....");  
            }  
            return false;  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return false;  
    }  
  
    /** 
     * 获取指定工作簿行数 
     *  
     * @param excelPath 
     * @param sheetName 
     * @return 
     */  
    public static int getExcelSheetRowNum(String excelPath, String sheetName) {  
        try {  
            if (FileUtil.checkFile(excelPath)) {  
                HSSFWorkbook wb = getHSSFWorkbook(excelPath);  
                if (ExcelUtils.checkSheet(wb, sheetName)) {  
                    HSSFSheet sheet = wb.getSheet(sheetName);  
                    return sheet.getLastRowNum() + 1;  
                } else {  
                    System.out.println("getExcelSheetRowNum:" + excelPath + "的"  
                            + sheetName + ":工作表不存在.....");  
                }  
  
            } else {  
                System.out.println(excelPath + ":文件不存在.....");  
            }  
            return 0;  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return 0;  
    }  
  
    /** 
     * 删除行 
     *  
     * @excelPath 
     * @param sheetName 
     * @param row 
     * @return 
     */  
    public static boolean deleteRow(String excelPath, String sheetName, int row) {  
        if (getExcelSheetRowNum(excelPath, sheetName) > row) {  
            try {  
                if (FileUtil.checkFile(excelPath)) {  
                    HSSFWorkbook wb = getHSSFWorkbook(excelPath);  
                    if (ExcelUtils.checkSheet(wb, sheetName)) {  
                        HSSFSheet sheet = wb.getSheet(sheetName);  
                        if (sheet.getRow(row) != null)  
                            sheet.removeRow(sheet.getRow(row));  
                        return outputHSSFWorkbook(wb, excelPath);  
                    } else {  
                        System.out.println("deleteRow:" + excelPath + "的"  
                                + sheetName + ":工作表不存在.....");  
                    }  
  
                } else {  
                    System.out.println(excelPath + ":文件不存在.....");  
                }  
                return false;  
            } catch (Exception e) {  
                e.printStackTrace();  
            }  
        } else {  
            System.out.println("不存在指定行");  
            return true;  
        }  
        return false;  
    }  
  
    /** 
     * 获取所有数据 
     *  
     * @excelPath 
     * @param sheetName 
     * @return 
     */  
    public static List<List> getAllData(String excelPath, String sheetName) {  
        try {  
            List<List> list = new ArrayList<List>();  
            if (FileUtil.checkFile(excelPath)) {  
                HSSFWorkbook wb = getHSSFWorkbook(excelPath);  
                if (ExcelUtils.checkSheet(wb, sheetName)) {  
                    HSSFSheet sheet = wb.getSheet(sheetName);  
                    for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {  
                        if (sheet.getRow(i) != null) {  
                            List<String> rowList = new ArrayList<String>();  
                            HSSFRow aRow = sheet.getRow(i);  
                            for (int cellNumOfRow = 0; cellNumOfRow < aRow  
                                    .getLastCellNum(); cellNumOfRow++) {  
                                if (null != aRow.getCell(cellNumOfRow)) {  
                                    HSSFCell aCell = aRow.getCell(cellNumOfRow);// 获得列值  
                                    if (aCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {  
                                        rowList.add(aCell.getStringCellValue());  
                                    } else if (aCell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {  
                                        rowList.add(String.valueOf(  
                                                aCell.getNumericCellValue())  
                                                .replace(".0", ""));  
                                    }  
                                } else {  
                                    rowList.add("");  
                                }  
                            }  
                            list.add(rowList);  
                        }  
                    }  
                    return list;  
                } else {  
                    System.out.println("getAllData:" + excelPath + "的"  
                            + sheetName + ":工作表不存在.....");  
                }  
            } else {  
                System.out.println(excelPath + ":文件不存在.....");  
            }  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return null;  
    }  
  
    /** 
     * 读取部分数据 
     *  
     * @param excelPath 
     * @param sheetName 
     * @param start 
     * @param end 
     * @return 
     */  
    public static List<List> getDatas(String excelPath, String sheetName,  
            int start, int end) {  
        try {  
            List<List> list = new ArrayList<List>();  
            if (FileUtil.checkFile(excelPath)) {  
                HSSFWorkbook wb = getHSSFWorkbook(excelPath);  
                if (ExcelUtils.checkSheet(wb, sheetName)) {  
                    HSSFSheet sheet = wb.getSheet(sheetName);  
                    for (int i = start - 1; i < end; i++) {  
                        if (sheet.getRow(i) != null) {  
                            List<String> rowList = new ArrayList<String>();  
                            HSSFRow aRow = sheet.getRow(i);  
                            for (int cellNumOfRow = 0; cellNumOfRow < aRow  
                                    .getLastCellNum(); cellNumOfRow++) {  
                                if (null != aRow.getCell(cellNumOfRow)) {  
                                    HSSFCell aCell = aRow.getCell(cellNumOfRow);// 获得列值  
                                    if (aCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {  
                                        rowList.add(aCell.getStringCellValue());  
                                    } else if (aCell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {  
                                        rowList.add(String.valueOf(  
                                                aCell.getNumericCellValue())  
                                                .replace(".0", ""));  
                                    }  
                                } else {  
                                    rowList.add("");  
                                }  
                            }  
                            list.add(rowList);  
                        }  
                    }  
                    return list;  
                } else {  
                    System.out.println("getDatas:" + excelPath + "的"  
                            + sheetName + ":工作表不存在.....");  
                }  
            } else {  
                System.out.println(excelPath + ":文件不存在.....");  
            }  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return null;  
    }  
  
    /** 
     * 获取指定行列数据 
     *  
     * @param excelPath 
     * @param sheetName 
     * @param row 
     * @param cellNum 
     * @return 
     */  
    public static String getData(String excelPath, String sheetName, int row,  
            int cellNum) {  
        if (getExcelSheetRowNum(excelPath, sheetName) > row) {  
            try {  
                if (FileUtil.checkFile(excelPath)) {  
                    HSSFWorkbook wb = getHSSFWorkbook(excelPath);  
                    if (ExcelUtils.checkSheet(wb, sheetName)) {  
                        HSSFSheet sheet = wb.getSheet(sheetName);  
                        HSSFCell aCell = sheet.getRow(row).getCell(cellNum);  
                        if (aCell != null)  
                            return aCell.getStringCellValue();  
                    } else {  
                        System.out.println("getData" + excelPath + "的"  
                                + sheetName + ":工作表不存在.....");  
                    }  
                } else {  
                    System.out.println(excelPath + ":文件不存在.....");  
                }  
            } catch (Exception e) {  
                e.printStackTrace();  
            }  
        }  
        return null;  
    }  
  
    /**
     * 取得指定单元格行和列
     * @param keyMap 所有单元格行、列集合
     * @param key 单元格标识
     * @return 0：列 1：行（列表型数据不记行，即1无值）
     */
    public static int[] getPos(HashMap keyMap, String key){
        int[] ret = new int[0];
          
        String val = (String)keyMap.get(key);
          
        if(val == null || val.length() == 0)
            return ret;
          
        String pos[] = val.split(",");
          
        if(pos.length == 1 || pos.length == 2){
            ret = new int[pos.length];
            for(int i0 = 0; i0 < pos.length; i0++){
                if(pos[i0] != null && pos[i0].trim().length() > 0){
                    ret[i0] = Integer.parseInt(pos[i0].trim());
                } else {
                    ret[i0] = 0;
                }
            }
        }
        return ret;
    }
      
    /**
     * 取对应格子的值
     * @param sheet
     * @param rowNo 行
     * @param cellNo 列
     * @return
     * @throws IOException
     */
    public static String getCellValue(Sheet sheet,int rowNo,int cellNo) {
        String cellValue = null;
        Row row = sheet.getRow(rowNo);
        Cell cell = row.getCell(cellNo);
        if (cell != null) {
            if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                DecimalFormat df = new DecimalFormat("0");
                cellValue = getCutDotStr(df.format(cell.getNumericCellValue()));
            } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                cellValue = cell.getStringCellValue();
            }
            if (cellValue != null) {
                cellValue = cellValue.trim();
            }           
        } else {
            cellValue = null;
        }
        return cellValue;
    }
       
    /**
     * 取整数
     * @param srcString
     * @return
     */
    private static String getCutDotStr(String srcString) {
        String newString = "";
        if (srcString != null && srcString.endsWith(".0")) {
            newString = srcString.substring(0,srcString.length()-2);
        } else {
            newString = srcString;
        }
        return newString;
    }   
      
    /**
     * 读数据模板
     * @param 模板地址
     * @throws IOException
     */
    public static HashMap[] getTemplateFile(String templateFileName) throws IOException {    
        FileInputStream fis = new FileInputStream(templateFileName);
         
        Workbook wbPartModule = null;
        if(templateFileName.endsWith(".xlsx")){
            wbPartModule = new XSSFWorkbook(fis);
        }else if(templateFileName.endsWith(".xls")){
            wbPartModule = new HSSFWorkbook(fis);
        }
        int numOfSheet = wbPartModule.getNumberOfSheets();
        HashMap[] templateMap = new HashMap[numOfSheet];
        for(int i = 0; i < numOfSheet; i++){
            Sheet sheet = wbPartModule.getSheetAt(i);
            templateMap[i] = new HashMap();
            readSheet(templateMap[i], sheet);
        }
        fis.close();
        return templateMap;
    }
    
    /**
     * 读数据模板
     * @param 模板地址
     * @throws IOException
     */
    public static HashMap[] getJsonTemplateFile(String templateFileName) throws IOException {    
        FileInputStream fis = new FileInputStream(templateFileName);
         
        Workbook wbPartModule = null;
        if(templateFileName.endsWith(".xlsx")){
            wbPartModule = new XSSFWorkbook(fis);
        }else if(templateFileName.endsWith(".xls")){
            wbPartModule = new HSSFWorkbook(fis);
        }
        int numOfSheet = wbPartModule.getNumberOfSheets();
        HashMap[] templateMap = new HashMap[numOfSheet];
        for(int i = 0; i < numOfSheet; i++){
            Sheet sheet = wbPartModule.getSheetAt(i);
            templateMap[i] = new HashMap();
            readSheetJsonTemp2(templateMap[i], sheet);
        }
        fis.close();
        return templateMap;
    }
      
    /**
     * 读模板数据的样式值置等信息
     * @param keyMap
     * @param sheet
     */
    private static void readSheet(HashMap keyMap, Sheet sheet){
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
          
        for (int j = firstRowNum; j <= lastRowNum; j++) {
            Row rowIn = sheet.getRow(j);
            if(rowIn == null) {
                continue;
            }
            int firstCellNum = rowIn.getFirstCellNum();
            int lastCellNum = rowIn.getLastCellNum();
            if(lastCellNum <=0) {
            	continue;
            }
            for (int k = firstCellNum; k <= lastCellNum; k++) {
//              Cell cellIn = rowIn.getCell((short) k);
                Cell cellIn = rowIn.getCell(k);
                if(cellIn == null) {
                    continue;
                }
                  
                int cellType = cellIn.getCellType();
                if(Cell.CELL_TYPE_STRING != cellType) {
                    continue;
                }
                String cellValue = cellIn.getStringCellValue();
                if(cellValue == null) {
                    continue;
                }
                cellValue = cellValue.trim();
                if(cellValue.length() > 1 && cellValue.substring(0,1).equals(":")) {
                    String key = cellValue.substring(1, cellValue.length());
                    String keyPos = Integer.toString(k)+","+Integer.toString(j);
                    keyMap.put(key, keyPos);
                    keyMap.put(key+"CellStyle$", cellIn.getCellStyle());
                } else if(cellValue.length() > 2 && cellValue.substring(0,2).equals("<%")) {
                    String key = cellValue.substring(2, cellValue.length());
                    String keyPos = Integer.toString(k)+","+Integer.toString(j);
                    keyMap.put(key, keyPos);
                    keyMap.put(key+"CellStyle$", cellIn.getCellStyle());
                } else if(cellValue.length() > 3 && cellValue.substring(0,3).equals("<!%")) {
                    String key = cellValue.substring(3, cellValue.length());
                    keyMap.put("STARTCELL", Integer.toString(j));
                    keyMap.put(key, Integer.toString(k));
                    keyMap.put(key+"CellStyle$", cellIn.getCellStyle());
                }
            }
        }
    }
      
    /**
     * 读模板数据的样式值置等信息
     * @param keyMap 存放模板变量值位置、样式
     * @param sheet  工作表对象
     */
    public static void readSheetJsonTemp(HashMap keyMap, Sheet sheet){
    	
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        String loopFlag = "s"; // 循环的标识；s-单个值；l-标识循环
          
        for (int j = firstRowNum; j <= lastRowNum; j++) {
        	loopFlag = "s";
        	// 将每行key放到一个map中
        	HashMap rowMap =  new HashMap();
            Row rowIn = sheet.getRow(j);
            if(rowIn == null) {
                continue;
            }
            int firstCellNum = rowIn.getFirstCellNum();
            int lastCellNum = rowIn.getLastCellNum();
            if(lastCellNum <=0) {
            	continue;
            }
            for (int k = firstCellNum; k <= lastCellNum; k++) {
//              Cell cellIn = rowIn.getCell((short) k);
                Cell cellIn = rowIn.getCell(k);
                if(cellIn == null) {
                    continue;
                }
                  
                int cellType = cellIn.getCellType();
                if(Cell.CELL_TYPE_STRING != cellType) {
                    continue;
                }
                String cellValue = cellIn.getStringCellValue();
                if(cellValue == null) {
                    continue;
                }
                cellValue = cellValue.trim();
                if(cellValue.length() > 1 && cellValue.substring(0,1).equals(":")) {
                    String key = cellValue.substring(1, cellValue.length());
                    String keyPos = Integer.toString(k)+","+Integer.toString(j);
                    rowMap.put(key, keyPos);
                    rowMap.put(key+"CellStyle$", cellIn.getCellStyle());
                    // 如果存在标识符[n]，说明该行有值要进行循环
                    if(key.contains("[n]")) {
                    	loopFlag = "l";
                    }
                } 
            }
            if(rowMap.size() > 0) {
	            // 将行map存放到keymap中
	            keyMap.put(loopFlag+j,rowMap);
            }
        }
    }
    
    /**
     * 读模板数据的样式值置等信息
     * @param keyMap 存放模板变量值位置、样式
     * @param sheet  工作表对象
     */
    public static void readSheetJsonTemp2(HashMap keyMap, Sheet sheet){
    	
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        List<HashMap> singleList = new ArrayList<HashMap>(); // 存放当前行完全不需要循环的数据map
        List<HashMap> loopList = new ArrayList<HashMap>();   // 存放当前行有数据需要循环的map
        String loopFlag = "s"; // 循环的标识；s-单个值；l-标识循环
          
        for (int j = firstRowNum; j <= lastRowNum; j++) {
        	loopFlag = "s";
        	// 将每行key放到一个map中
        	HashMap rowMap =  new HashMap();
            Row rowIn = sheet.getRow(j);
            if(rowIn == null) {
                continue;
            }
            int firstCellNum = rowIn.getFirstCellNum();
            int lastCellNum = rowIn.getLastCellNum();
            if(lastCellNum <=0) {
            	continue;
            }
            for (int k = firstCellNum; k <= lastCellNum; k++) {
//              Cell cellIn = rowIn.getCell((short) k);
                Cell cellIn = rowIn.getCell(k);
                if(cellIn == null) {
                    continue;
                }
                  
                int cellType = cellIn.getCellType();
                if(Cell.CELL_TYPE_STRING != cellType) {
                    continue;
                }
                String cellValue = cellIn.getStringCellValue();
                if(cellValue == null) {
                    continue;
                }
                cellValue = cellValue.trim();
                if(cellValue.length() > 1 && cellValue.substring(0,1).equals(":")) {
                    String key = cellValue.substring(1, cellValue.length());
                    String keyPos = Integer.toString(k)+","+Integer.toString(j);
                    rowMap.put(key, keyPos);
                    rowMap.put(key+"CellStyle$", cellIn.getCellStyle());
                    // 如果存在标识符[n]，说明该行有值要进行循环
                    if(key.contains("[n]")) {
                    	loopFlag = "l";
                    }
                } 
            }
            // 根据得到的循环标识，进行封装
            if("s".equals(loopFlag)) {
            	singleList.add(rowMap);
            } else {
            	loopList.add(rowMap);
            }
            
        }
        
        // 将行list存放到keymap中
        if(singleList.size() > 0) {
            
            keyMap.put("s",singleList);
        }
        if(loopList.size() > 0) {
            // 将行map存放到keymap中
            keyMap.put("l",loopList);
        }
    }
    /**
     * 获取格式，不适于循环方法中使用，wb.createCellStyle()次数超过4000将抛异常
     * @param keyMap
     * @param key
     * @return
     */
    public static CellStyle getStyle(HashMap keyMap, String key,Workbook wb) {
        CellStyle cellStyle = null;      
          
        cellStyle = (CellStyle) keyMap.get(key+"CellStyle$");
        //当字符超出时换行
        cellStyle.setWrapText(true);
        CellStyle newStyle = wb.createCellStyle();
        newStyle.cloneStyleFrom(cellStyle);
        return newStyle;
    }
    /**
     * Excel单元格输出
     * @param sheet
     * @param row 行
     * @param cell 列
     * @param value 值
     * @param cellStyle 样式
     */
    public static void setValue(Sheet sheet, int row, int cell, Object value, CellStyle cellStyle){
        Row rowIn = sheet.getRow(row);
        if(rowIn == null) {
            rowIn = sheet.createRow(row);
        }
        Cell cellIn = rowIn.getCell(cell);
        if(cellIn == null) {
            cellIn = rowIn.createCell(cell);
        }
        if(cellStyle != null) {  
            //修复产生多超过4000 cellStyle 异常
            //CellStyle newStyle = wb.createCellStyle();
            //newStyle.cloneStyleFrom(cellStyle);
            cellIn.setCellStyle(cellStyle);
        }
        //对时间格式进行单独处理
        if(value==null){
            cellIn.setCellValue("");
        }else{
            if (isCellDateFormatted(cellStyle)) {
                cellIn.setCellValue((Date) value);
            } else {
                cellIn.setCellValue(new XSSFRichTextString(value.toString()));
            }
        }
    }
      
    /**
     * 根据表格样式判断是否为日期格式
     * @param cellStyle
     * @return
     */
    public static boolean isCellDateFormatted(CellStyle cellStyle){
        if(cellStyle==null){
            return false;
        }
        int i = cellStyle.getDataFormat();
        String f = cellStyle.getDataFormatString();
         
        return org.apache.poi.ss.usermodel.DateUtil.isADateFormat(i, f);
    }
    /**
     * 适用于导出的数据Excel格式样式重复性较少
     * 不适用于循环方法中使用
     * @param wbModule
     * @param sheet
     * @param pos 模板文件信息
     * @param startCell 开始的行
     * @param value 要填充的数据
     * @param cellStyle 表格样式
     */
    public static void createCell(Workbook wbModule, Sheet sheet,HashMap pos, int startCell,Object value,String cellStyle){
        int[] excelPos = getPos(pos, cellStyle);
        setValue(sheet, startCell, excelPos[0], value, getStyle(pos, cellStyle,wbModule));
    }
    
    /** 
     * 找到需要插入的行数，并新建一个POI的row对象 
     * @param sheet 
     * @param rowIndex 
     * @return 
     */  
    public static Row createRow(Sheet sheet, Integer rowIndex) {  
         Row row = null;  
         if (sheet.getRow(rowIndex) != null) {  
             int lastRowNo = sheet.getLastRowNum();  
             sheet.shiftRows(rowIndex, lastRowNo, 1);  
         }  
         row = sheet.createRow(rowIndex);  
         return row;  
     }  
    /************************************XSSF***************************************
    /** 
     * @param args 
     */  
    public static void main(String[] args) {  
        // 创建新的excle  
        // ExcelUtils.createExcelFile("D:\\test.xls");  
  
        // 插入新的工作表  
        // ExcelUtils.insertSheet("D:\\test.xls", "333");  
  
        // 检查是否存在工作表  
        // HSSFWorkbook wb;  
        // try {  
        // wb = getHSSFWorkbook("D:\\test.xls");  
        // System.out.println(ExcelUtils.checkSheet(wb, "3334"));  
        // } catch (Exception e) {  
        // // TODO Auto-generated catch block  
        // e.printStackTrace();  
        // }  
  
        // 插入一行数据  
        // ExcelUtils.insertOrUpadateRowDatas("D:\\test.xls", "333", 5, "123",  
        // "0000");  
  
        // 插入单元格数据  
        // ExcelUtils.insertOrUpdateCell("D:\\test.xls", "123", 5, 1, "0000");  
  
        // 删除指定工作表所有的行数据、  
        // ExcelUtils.cleanExcelFile("D:\\test.xls", "123");  
  
        // 删除指定行  
        // ExcelUtils.deleteRow("D:\\test.xls", "123", 0);  
  
        // 获取所有的数据  
        // List<List> list = ExcelUtils.getAllData("D:\\test.xls", "333");  
        // for (int i = 0; i < list.size(); i++) {  
        // List<String> rowList = list.get(i);  
        // for (int j = 0; j < rowList.size(); j++) {  
        // System.out.println(rowList.get(j));  
        // }  
        // }  
  
        // 获取部分的数据  
        // List<List> list = ExcelUtils.getDatas(  
        // "C:\\Users\\Ken\\Desktop\\测试名单.xls", "Sheet1", 5, 374);  
        // for (int i = 0; i < list.size(); i++) {  
        // List<String> rowList = list.get(i);  
        // for (int j = 0; j < rowList.size(); j++) {  
        // System.out.print(rowList.get(j) + " ");  
        // }  
        // System.out.println();  
        // }  
  
        // 获取指定行列数据  
        // System.out.println(ExcelUtils.getData("D:\\test.xls", "333", 5, 1));  
  
        // 复制指定工作表  
        ExcelUtils.copySheet("D:\\data.xlsx", "333", 0);  
  
    }  
}  