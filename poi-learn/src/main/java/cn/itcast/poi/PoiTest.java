package cn.itcast.poi;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;

public class PoiTest {

    @Test
    public void excelRead()throws Exception{
        FileInputStream in = new FileInputStream(new File("C:/Users/asus/Desktop/求职资料/全国各城市各学科缺口数分析图(2019.3.13).xls"));
        //得到POI文件系统对象
        POIFSFileSystem fs = new POIFSFileSystem(in);
        //得到Excel工作簿对象
        HSSFWorkbook wk = new HSSFWorkbook(fs);
        //得到Excel工作簿的第一页，即excel工作表对象
        HSSFSheet sheet = wk.getSheetAt(0);
        //遍历工作表
        //遍历行对象
        for (Row row : sheet) {
            //打印行索引
            //System.out.println(row.getRowNum());
            //遍历单元格对象
            //表头不要打印
            for (Cell cell : row) {
                //获取每个单元格的类型
                int cellType = cell.getCellType();
                if(cellType == cell.CELL_TYPE_BLANK){
                    //System.out.println("空格类型");
                    System.out.print("\t");
                }
                if(cellType == cell.CELL_TYPE_NUMERIC){
                    //System.out.println("数字类型");
                    System.out.print(cell.getNumericCellValue()+"\t");
                }
                if(cellType == cell.CELL_TYPE_STRING){
                    //System.out.println("字符串类型");
                    System.out.print(cell.getStringCellValue()+"\t");
                }
            }
            //换行
            System.out.println();
        }


    }
}
