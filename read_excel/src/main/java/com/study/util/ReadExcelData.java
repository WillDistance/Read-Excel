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
        List listData = readExcelToList("C:\\Users\\Thinkive\\Desktop\\����֤ȯͶ�̸���\\��Ŀ��Ϣ��¼\\��Ŀģ��.xlsx");
        for(int i=1;i<listData.size();i++){
            List row = (List) listData.get(i);
            System.out.println(row.toString());
        }
        
        List<SubjectEntity> subjects = new ArrayList<SubjectEntity>();
        String preRow = "";//��¼��һ�е����
        int subjectRank = 0; //��¼��Ŀ���� ��û�и�������ʱʹ��
        SubjectEntity se = null;//��Ŀʵ��
        for(int i=1;i<listData.size();i++){
            List row = (List) listData.get(i); //��ȡһ�е����ݣ�ÿһ�е����ݰ�˳��洢��list��
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
                se.setRank(rank); //������Ŀ����
                se.setContent((String) row.get(1)); //����������Ŀ
                se.setType((String) row.get(2)); //������Ŀ���ͣ�0:��ѡ��1����ѡ
                se.setIs_myd((String) row.get(3)); //�����Ƿ��������Ŀ ;0:��1����
                se.setIs_must((String) row.get(4)); //�����Ƿ�Ϊ�ش��⣻0:�Ǳش��⣬1���ش���
                List<String> option = new ArrayList<String>();
                //��ȡ5~8�е�ѡ������
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
                    se = new SubjectEntity();//���´�����Ŀ����
                    subjectRank++;
                    preRow = (String) row.get(1);
                    String rank = (String) row.get(0);
                    if(StringHelper.isBlank(rank))
                    {
                        rank = subjectRank+"";
                    }
                    se.setRank(rank); //������Ŀ����
                    se.setContent((String) row.get(1)); //����������Ŀ
                    se.setType((String) row.get(2)); //������Ŀ���ͣ�0:��ѡ��1����ѡ
                    se.setIs_myd((String) row.get(3)); //�����Ƿ��������Ŀ ;0:��1����
                    se.setIs_must((String) row.get(4)); //�����Ƿ�Ϊ�ش��⣻0:�Ǳش��⣬1���ش���
                }
                
                List<String> option = new ArrayList<String>();
                //��ȡ5~8�е�ѡ������
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
            logger.error("�ļ�������");
            return null;
        }
        InputStream inputStream = null;
        ArrayList sheetList = new ArrayList();
        try
        {
            // ������Excel�������ļ�������
            inputStream = new FileInputStream(file);
            Workbook workbook = createWorkbook(inputStream);
            
            Sheet sheet = workbook.getSheetAt(0);//��ȡ��һ��������
            
            int start = sheet.getFirstRowNum();//��ȡ��ʼ�ͽ������к�
            int end = sheet.getLastRowNum();
            
            Row row_0 = sheet.getRow(0);
            if(row_0 == null)
            {
                logger.error("δ��ȡ��������");
                return null;
            }
            short startCell = row_0.getFirstCellNum();//��ȡ��ʵ�����ݵĵ�1����Ԫ��0�ǵ�һ��
            short endCell = row_0.getLastCellNum();//��ȡ��ʵ�����ݵ����1����Ԫ��1�ǵ�һ��
            
            for (int i = start; i <= end; i++)
            {
                Row row = sheet.getRow(i);//��ȡһ�е�����
                if ( row == null )
                {
                    continue;
                }
                ArrayList rowList = new ArrayList();
                //��¼
                for (int j = startCell; j < endCell; j++)
                {
                    Cell cell = row.getCell(j);
                    if(isMergedRegion(sheet, i, j))//�жϵ�ǰ��Ԫ���Ƿ��Ǻϲ���Ԫ��
                    {
                        rowList.add(getMergedRegionValue(sheet, i, j));//��ȡ�ϲ���Ԫ���ֵ
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
            logger.error("��ȡ�ļ�����", ex);
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
                    logger.error("�ر�����������", ex);
                }
            }
        }
        return null;
    }
    
    
    /**
     * ��ȡ��Ԫ���ֵ
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
     * ��ȡ�ϲ���Ԫ���ֵ
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
     * �ж�ָ���ĵ�Ԫ���Ƿ��Ǻϲ���Ԫ��
     * @param sheet
     * @param row ���±�
     * @param column ���±�
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
        
        throw new IllegalArgumentException("��excel�汾Ŀǰpoi��������");
    }
}
