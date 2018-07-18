package com.heyu.excelParser;

import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

public class App
{
    private static final int INDEX_ZERO = 0;
    private static final int INDEX_ONE = 0;

    public static void main(String[] agrs) throws IOException
    {
        String workPath = System.getProperty("user.dir");
        String excelPath = workPath + "\\excelDemo\\KPI-2007.xlsx";
        getExcelContent(excelPath);
    }

    /**
     * 获取excel里面的KPI内容
     * @param excelPath excel路径
     * @throws IOException IO异常
     */
    private static void getExcelContent(String excelPath) throws IOException
    {
        //获取excel
        XSSFWorkbook excel = getExcelObject(excelPath);
        if (excel == null)
        {
            return;
        }
        int sheets = excel.getNumberOfSheets();
        //遍历第一个sheet
        XSSFSheet sheet = excel.getSheetAt(INDEX_ZERO);

        //获取表头
        XSSFRow rowTitle = sheet.getRow(INDEX_ZERO);
        Map<Integer, String> category = getExcelTitle(rowTitle);

        //根据表头遍历每一列数据
        getExcelData(category, sheet);

        closeExcel(excel);
    }

    /**
     * 获取excel对象
     * @param excelPath
     * @return
     * @throws IOException
     */
    private static XSSFWorkbook getExcelObject(String excelPath) throws IOException
    {
        File file = new File(excelPath);
        InputStream in = new FileInputStream(file);
        //得到整个excel对象
        XSSFWorkbook excel = new XSSFWorkbook(in);

        return excel;
    }

    /**
     * 获取excel里面的数据
     * @param category
     * @param sheet
     */
    private static void getExcelData(Map<Integer, String> category, XSSFSheet sheet)
    {
        if (MapUtils.isEmpty(category) || sheet == null)
        {
            return;
        }

        for (Map.Entry<Integer, String> entry : category.entrySet())
        {
            int colTitle = entry.getKey();
            String title = entry.getValue();
            System.out.println();

            //遍历每一行指定的列数据
            for (int rowNum = INDEX_ONE; rowNum <= sheet.getLastRowNum(); ++rowNum)
            {
                XSSFRow row = sheet.getRow(rowNum);
                if (row == null)
                {
                    continue;
                }
                XSSFCell cell = row.getCell(colTitle);
                if (cell == null)
                {
                    continue;
                }
                String cellValue = cell.toString();
                if (StringUtils.isEmpty(cellValue))
                {
                    continue;
                }
                System.out.println("Title: " + title + ", value: " + cellValue);
            }
        }
    }

    /**
     * 获取Excel中的title
     * @param rowTitle
     * @return
     */
    private static Map<Integer, String> getExcelTitle(XSSFRow rowTitle)
    {
        if (rowTitle == null)
        {
            return null;
        }

        Map<Integer, String> excelTitle = new HashMap<Integer, String>();
        for (int i = 0; i < rowTitle.getLastCellNum(); ++i)
        {
            XSSFCell cell = rowTitle.getCell(i);
            String cellStr = cell.toString();
            int colIndex = cell.getColumnIndex();
            excelTitle.put(colIndex, colIndex + "-" + cellStr);
        }
        System.out.println();

        return excelTitle;
    }

    /**
     * 释放excel资源
     *
     * @param excel
     * @throws IOException
     */
    private static void closeExcel(XSSFWorkbook excel) throws IOException
    {
        if (excel != null)
        {
            excel.close();
        }
    }
}
