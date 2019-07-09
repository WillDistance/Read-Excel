package com.study.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.thinkive.base.util.StringHelper;

public class ReadExcelData
{
    private static Logger logger = Logger.getLogger(ReadExcelData.class);
    
    public static void main(String[] args)
    {
        List listData = readExcelToList("C:\\Users\\Thinkive\\Desktop\\华龙证券投教更新\\项目信息记录\\题目模板.xlsx");
        for(int i=1;i<listData.size();i++){
            List row = (List) listData.get(i);
            System.out.println(row.toString());
        }
        
        List<SubjectEntity> subjects = new ArrayList<SubjectEntity>();
        String preRow = "";//记录上一行的题干
        int subjectRank = 0; //记录题目排序 在没有给定排序时使用
        SubjectEntity se = null;//题目实体
        for(int i=1;i<listData.size();i++){
            List row = (List) listData.get(i); //获取一行的数据，每一行的数据按顺序存储在list中
            if(StringHelper.isBlank(preRow))
            {
                subjectRank++;
                preRow = (String) row.get(1);
                se = new SubjectEntity();
                String rank = (String) row.get(0);
                if(StringHelper.isBlank(rank))
                {
                    rank = subjectRank+"";
                }
                se.setRank(rank); //设置题目排序
                se.setContent((String) row.get(1)); //设置问题题目
                se.setType((String) row.get(2)); //设置题目类型；0:单选，1、多选
                se.setIs_myd((String) row.get(3)); //设置是否满意度题目 ;0:否，1、是
                se.setIs_must((String) row.get(4)); //设置是否为必答题；0:非必答题，1、必答题
                List<String> option = new ArrayList<String>();
                //获取5~8列的选项数据
                for (int j = 5; j < row.size(); j++)
                {
                    if(j==5&&StringHelper.isBlank((String) row.get(j)))
                    {
                        option.add(j+"");
                    }
                    else
                    {
                        option.add((String) row.get(j));
                    }
                }
                se.addOptions(option);
            }
            else
            {
                if(!preRow.equals((String) row.get(1)))
                {
                    subjects.add(se);
                    se = new SubjectEntity();//重新创建题目对象
                    subjectRank++;
                    preRow = (String) row.get(1);
                    String rank = (String) row.get(0);
                    if(StringHelper.isBlank(rank))
                    {
                        rank = subjectRank+"";
                    }
                    se.setRank(rank); //设置题目排序
                    se.setContent((String) row.get(1)); //设置问题题目
                    se.setType((String) row.get(2)); //设置题目类型；0:单选，1、多选
                    se.setIs_myd((String) row.get(3)); //设置是否满意度题目 ;0:否，1、是
                    se.setIs_must((String) row.get(4)); //设置是否为必答题；0:非必答题，1、必答题
                }
                
                List<String> option = new ArrayList<String>();
                //获取5~8列的选项数据
                for (int j = 5; j < row.size(); j++)
                {
                    if(j==5&&StringHelper.isBlank((String) row.get(j)))
                    {
                        option.add(j+"");
                    }
                    else
                    {
                        option.add((String) row.get(j));
                    }
                }
                se.addOptions(option);
            }
        }
        subjects.add(se);
        
        for (SubjectEntity subjectEntity : subjects)
        {
            System.out.println(subjectEntity.toString());
        }
    }
    
    @SuppressWarnings({ "unchecked", "rawtypes" })
    public static List readExcelToList(String filePath)
    {
        File file = new File(filePath);
        if(!file.exists() || !file.isFile())
        {
            logger.error("文件不存在");
            return null;
        }
        InputStream inputStream = null;
        ArrayList sheetList = new ArrayList();
        try
        {
            // 创建对Excel工作簿文件的引用
            inputStream = new FileInputStream(file);
            Workbook workbook = createWorkbook(inputStream);
            
            Sheet sheet = workbook.getSheetAt(0);//获取第一个工作表
            
            int start = sheet.getFirstRowNum();//获取开始和结束的行号
            int end = sheet.getLastRowNum();
            
            Row row_0 = sheet.getRow(0);
            if(row_0 == null)
            {
                logger.error("未获取到标题行");
                return null;
            }
            short startCell = row_0.getFirstCellNum();//获取真实有数据的第1个单元格，0是第一个
            short endCell = row_0.getLastCellNum();//获取真实有数据的最后1个单元格，1是第一个
            
            for (int i = start; i <= end; i++)
            {
                Row row = sheet.getRow(i);//获取一行的数据
                if ( row == null )
                {
                    continue;
                }
                ArrayList rowList = new ArrayList();
                //记录
                for (int j = startCell; j < endCell; j++)
                {
                    Cell cell = row.getCell(j);
                    if(isMergedRegion(sheet, i, j))//判断当前单元格是否是合并单元格
                    {
                        rowList.add(getMergedRegionValue(sheet, i, j));//读取合并单元格的值
                    }
                    else
                    {
                        rowList.add(getCellValue(cell));
                    }
                }
                sheetList.add(rowList);
            }
            return sheetList;
        }
        catch (Exception ex)
        {
            logger.error("读取文件出错", ex);
        }
        finally
        {
            if ( inputStream != null )
            {
                try
                {
                    inputStream.close();
                    inputStream = null;
                }
                catch (Exception ex)
                {
                    logger.error("关闭输入流出错", ex);
                }
            }
        }
        return null;
    }
    
    
    /**
     * 获取单元格的值
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell){
        if(cell == null) return "";
        cell.setCellType(Cell.CELL_TYPE_STRING);
        String value = cell.getStringCellValue();
        if(StringHelper.isBlank(value)) return "";
        return value;
    }
    
    /**
     * 获取合并单元格的值
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static String getMergedRegionValue(Sheet sheet ,int row , int column){
        int sheetMergeCount = sheet.getNumMergedRegions();

        for(int i = 0 ; i < sheetMergeCount ; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell) ;
                }
            }
        }

        return null ;
    }


    /**
     * 判断指定的单元格是否是合并单元格
     * @param sheet
     * @param row 行下标
     * @param column 列下标
     * @return
     */
    public static boolean isMergedRegion(Sheet sheet,int row ,int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    return true;
                }
            }
        }
        return false;
    }
    
    
    
    private static Workbook createWorkbook(InputStream inputStream) throws InvalidFormatException, IOException
    {
        if ( !inputStream.markSupported() )
        {
            inputStream = new PushbackInputStream(inputStream, 8);
        }
        
        if ( POIFSFileSystem.hasPOIFSHeader(inputStream) )
        {
            return new HSSFWorkbook(inputStream);
        }
        else if ( POIXMLDocument.hasOOXMLHeader(inputStream) )
        {
            return new XSSFWorkbook(OPCPackage.open(inputStream));
        }
        
        throw new IllegalArgumentException("该excel版本目前poi解析不了");
    }
}
