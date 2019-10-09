package com.ames;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Demo {

    public static void main(String[] args) throws IOException {
        String path = "D:\\hlh";        //要遍历的路径
        File file = new File(path);        //获取其file对象
        File[] fs = file.listFiles();    //遍历path下的文件和目录，放在File数组中
        Map<String, Map<String, Integer>> typeAndRowCloumnMap = new HashMap();
        for (int i = 0; i < fs.length; i++) {                    //遍历File[]数组
            File[] files = fs[i].listFiles();
            for (File everyFile : files) {
                System.out.println(everyFile.getName());

                FileInputStream is = new FileInputStream(everyFile); //文件流
                Workbook workbook = new HSSFWorkbook(is); //这种方式 Excel 2003/2007/2010 都是可以处理的

                FormulaEvaluator evaluator= workbook.getCreationHelper().createFormulaEvaluator();
                //遍历每个Sheet
                Sheet sheet = workbook.getSheetAt(0);
                int rowCount = sheet.getPhysicalNumberOfRows(); //获取总行数
                Map<String, Integer> map;
                if (i == 0) {
                    map = new HashMap<>();
                } else {
                    map = typeAndRowCloumnMap.get(getTypeByFileName(everyFile));
                }
                //遍历每一行
                for (int r = 0; r < rowCount; r++) {
                    Row row = sheet.getRow(r);
                    int cellCount = 0; //获取总列数p
                    if(row==null){
                        continue;
                    }
                    try{
                        cellCount = row.getPhysicalNumberOfCells(); //获取总列数p
                    }catch (Exception e){
                        System.out.println(e);
                    }
                    //遍历每一列
                    List<Double> kg = new ArrayList();
                    for (int c = 0; c < cellCount; c++) {
                        Cell cell = row.getCell(c);
                        CellType cellType=null;
                        try{
                            cellType=cell.getCellTypeEnum();
                        }catch (Exception e){
                            continue;
                        }
                        if(CellType.NUMERIC==cellType){
                            Integer d = Integer.valueOf((int) cell.getNumericCellValue());
                            if(getTypeByFileName(everyFile).equals("血常规")&&c==3&&r==7){
                                System.out.println("zhuyi:"+everyFile.getName()+"zhi:"+d);
                            }
                            if (map.get("" + r +":"+ c) == null) {
                                map.put("" + r+":" + c, d);
                            } else {
                                try{
                                    map.put("" + r+":" + c, d + map.get("" + r +":"+ c));
                                }catch (Exception e){
                                    System.out.println(e);
                                }
                            }
                        }
                        if(CellType.FORMULA==cellType){
                            System.out.println("gogs:"+evaluator.evaluate(cell).getNumberValue());
                            Integer d =(int) evaluator.evaluate(cell).getNumberValue();
                            if (map.get("" + r +":"+ c) == null) {
                                map.put("" + r+":" + c, d);
                            } else {
                                try{
                                    map.put("" + r+":" + c, d + map.get("" + r +":"+ c));
                                }catch (Exception e){
                                    System.out.println(e);
                                }
                            }
                        }
                    }
                }
                try{
                    typeAndRowCloumnMap.put(getTypeByFileName(everyFile), map);
                }catch (Exception e){
                    System.out.println(e);
                }
            }
            System.out.println("统计了：" + fs[i].getName());
        }
        System.out.println("结果是：" + typeAndRowCloumnMap);

        String pathre = "D:\\hj";        //要遍历的路径
        File filere = new File(pathre);        //获取其file对象
        File[] fsre = filere.listFiles();
        for (File file1:fsre){
            FileInputStream is = new FileInputStream(file1); //文件流
            Workbook workbook = new HSSFWorkbook(is); //这种方式 Excel 2003/2007/2010 都是可以处理的
            //遍历每个Sheet
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getPhysicalNumberOfRows(); //获取总行数
            Map<String, Integer> map;
            //遍历每一行
            for (int r = 0; r < rowCount; r++) {
                Row row = sheet.getRow(r);
                int cellCount = 0; //获取总列数p
                if(row==null){
                    continue;
                }
                try{
                    cellCount = row.getPhysicalNumberOfCells(); //获取总列数p
                }catch (Exception e){
                    System.out.println(e);
                }
                //遍历每一列
                List<Double> kg = new ArrayList();
                for (int c = 0; c < cellCount; c++) {
                    if(typeAndRowCloumnMap.get(file1.getName().substring(0,file1.getName().indexOf("."))).get(""+r+":"+c)!=null){
                        Cell cell = row.getCell(c);
                        cell.setCellValue(typeAndRowCloumnMap.get(getTypeByFileName(file1)).get(""+r+":"+c));
                        cell.setCellType(CellType.NUMERIC);
                    }
                }
            }
            workbook.write(new FileOutputStream("D:\\hj\\"+file1.getName()));
        }


    }

    private static String  getTypeByFileName(File file) {

       return TypeAndCheckEnum.getTypeByFileName(file.getName());
    }


}
