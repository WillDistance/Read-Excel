package com.study.util;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelImportServiceImpl {

    public static void main(String[] args) throws FileNotFoundException, Exception
    {
        ExcelImportServiceImpl e = new ExcelImportServiceImpl();
        e.importExcel(new FileInputStream("C:\\Users\\Thinkive\\Desktop\\����֤ȯͶ�̸���\\��Ŀ��Ϣ��¼\\��Ŀģ��.xlsx"), "��Ŀģ��.xlsx");
    }
    
    public String importExcel(InputStream inputStream, String fileName) throws Exception{

        String message = "Import success";

        boolean isE2007 = false;
        //�ж��Ƿ���excel2007��ʽ
        if(fileName.endsWith("xlsx")){
            isE2007 = true;
        }

        int rowIndex = 0;
        try {
            InputStream input = inputStream;  //����������
            Workbook wb;
            //�����ļ���ʽ(2003����2007)����ʼ��
            if(isE2007){
                wb = new XSSFWorkbook(input);
            }else{
                wb = new HSSFWorkbook(input);
            }
            Sheet sheet = wb.getSheetAt(0);    //��õ�һ����
            int rowCount = sheet.getLastRowNum()+1;

            for(int i = 1; i < rowCount;i++){
                rowIndex = i;
                Row row ;

                for(int j = 0;j<8;j++){
                    if(isMergedRegion(sheet,i,j)){
                        System.out.print(getMergedRegionValue(sheet,i,j)+"\t");
                    }else{
                        row = sheet.getRow(i);
                        System.out.print(row.getCell(j)+"\t");
                    }
                }
                System.out.print("\n");
            }
        } catch (Exception ex) {
            ex.printStackTrace();
            message =  "Import failed, please check the data in "+rowIndex+" rows ";
        }
        return message;
    }


    /**
     * ��ȡ��Ԫ���ֵ
     * @param cell
     * @return
     */
    public  String getCellValue(Cell cell){
        if(cell == null) return "";
        return cell.getStringCellValue();
    }


    /**
     * �ϲ���Ԫ����,��ȡ�ϲ���
     * @param sheet
     * @return List<CellRangeAddress>
     */
    public  List<CellRangeAddress> getCombineCell(Sheet sheet)
    {
        List<CellRangeAddress> list = new ArrayList<CellRangeAddress>();
        //���һ�� sheet �кϲ���Ԫ�������
        int sheetmergerCount = sheet.getNumMergedRegions();
        //�������еĺϲ���Ԫ��
        for(int i = 0; i<sheetmergerCount;i++)
        {
            //��úϲ���Ԫ�񱣴��list��
            CellRangeAddress ca = sheet.getMergedRegion(i);
            list.add(ca);
        }
        return list;
    }

    private  int getRowNum(List<CellRangeAddress> listCombineCell,Cell cell,Sheet sheet){
        int xr = 0;
        int firstC = 0;
        int lastC = 0;
        int firstR = 0;
        int lastR = 0;
        for(CellRangeAddress ca:listCombineCell)
        {
            //��úϲ���Ԫ�����ʼ��, ������, ��ʼ��, ������
            firstC = ca.getFirstColumn();
            lastC = ca.getLastColumn();
            firstR = ca.getFirstRow();
            lastR = ca.getLastRow();
            if(cell.getRowIndex() >= firstR && cell.getRowIndex() <= lastR)
            {
                if(cell.getColumnIndex() >= firstC && cell.getColumnIndex() <= lastC)
                {
                    xr = lastR;
                }
            }

        }
        return xr;

    }
    /**
     * �жϵ�Ԫ���Ƿ�Ϊ�ϲ���Ԫ���ǵĻ��򽫵�Ԫ���ֵ����
     * @param listCombineCell ��źϲ���Ԫ���list
     * @param cell ��Ҫ�жϵĵ�Ԫ��
     * @param sheet sheet
     * @return
     */
    public  String isCombineCell(List<CellRangeAddress> listCombineCell,Cell cell,Sheet sheet)
            throws Exception{
        int firstC = 0;
        int lastC = 0;
        int firstR = 0;
        int lastR = 0;
        String cellValue = null;
        for(CellRangeAddress ca:listCombineCell)
        {
            //��úϲ���Ԫ�����ʼ��, ������, ��ʼ��, ������
            firstC = ca.getFirstColumn();
            lastC = ca.getLastColumn();
            firstR = ca.getFirstRow();
            lastR = ca.getLastRow();
            if(cell.getRowIndex() >= firstR && cell.getRowIndex() <= lastR)
            {
                if(cell.getColumnIndex() >= firstC && cell.getColumnIndex() <= lastC)
                {
                    Row fRow = sheet.getRow(firstR);
                    Cell fCell = fRow.getCell(firstC);
                    cellValue = getCellValue(fCell);
                    break;
                }
            }
            else
            {
                cellValue = "";
            }
        }
        return cellValue;
    }

    /**
     * ��ȡ�ϲ���Ԫ���ֵ
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public  String getMergedRegionValue(Sheet sheet ,int row , int column){
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
    private  boolean isMergedRegion(Sheet sheet,int row ,int column) {
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
    
}