package com.mo.codeDemo;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

public class TestPoi {
    public static void main(String[] args) {
        try {
            String path = Class.class.getClass().getResource("/").getPath();
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(path + "exportTemplate.xlsm")));

            //====================================begin=================0==========================
            //往sheet里面插入数据，然后通过POI绘制折线图
            /*workbook.setSheetHidden(workbook.getSheetIndex("template"),true);
            XSSFSheet sheet=workbook.createSheet("testSheet");
            sheet.getRow(10).getCell(3).setCellValue(5.2);
            Drawing drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1*17, 5, 1*17+16, 33);
            XSSFChart chart=sheet.createDrawingPatriarch().createChart(anchor);
            //List<XSSFChart> charts=sheet.createDrawingPatriarch().getCharts();
            //XSSFChart chart1 =charts.get(0);
            //创建绘图的类型   LineCahrtData 折线图
            LineChartData chartData = chart.getChartDataFactory().createLineChartData();
            //设置横坐标
            ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            //设置纵坐标
            ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
            leftAxis.setMajorTickMark(AxisTickMark.OUT);
            leftAxis.isVisible();
            //横坐标时间显示
            String[] timeArr = new String[]{"2017/7/1","2017/8/1","2017/9/1","2017/10/1","2017/11/1"};
            //纵坐标数据显示
            for (int i=0;i<10;i++){
                Double[] elements =new Double[]{i*1.11,i*1.22,i*1.33,i*1.44,i*1.55};
                //为chart的系列绑定数据源
                ChartDataSource<String> xAxis = DataSources.fromArray(timeArr);
                ChartDataSource<Double> dataAxis = DataSources.fromArray(elements);
                LineChartSeries chartSeries = chartData.addSeries(xAxis, dataAxis);
                chartSeries.setTitle("第"+i+"个点");
            }
            //开始绘制折线图
            chart.plot(chartData, bottomAxis, leftAxis);*/
            //以上这种方式通过POI绘制出来的图的样式无法控制，样式不符合要求的话无法修改
            //====================================end===========================================

            //================================begin==========================================
            //得到点的名字（用来给sheet进行命名）
            /*String[] pointNames=new String[]{"CX1","CX2"};
            for (int i = 0; i <pointNames.length ; i++) {
                //找到表格模板sheet，并克隆、填数据
                XSSFSheet templateSheet=null;
                XSSFSheet dataSheet=null;
                XSSFSheet exportSheet=null;
                Iterator<Sheet> sheetIterator=workbook.sheetIterator();
                while(sheetIterator.hasNext()){
                    XSSFSheet tempSheet=(XSSFSheet) sheetIterator.next();
                    if(tempSheet.getSheetName().equals("CX_template")){
                        templateSheet=tempSheet;
                    }
                    if(tempSheet.getSheetName().equals("CX_exportData")){
                        dataSheet=tempSheet;
                    }
                    if(tempSheet.getSheetName().equals("CX_export")){
                        exportSheet=tempSheet;
                    }
                }
                //克隆数据sheet
                XSSFSheet cloneDataSheet=workbook.cloneSheet(workbook.getSheetIndex(dataSheet));
                String name=cloneDataSheet.getSheetName();
                //修改名字，名字必须有一定规则（导出的sheet的名字+"Data"=数据的sheet的名字）,这样VBA代码才能够将图表指向对应的data所在的sheet
                workbook.setSheetName(workbook.getSheetIndex(name),pointNames[i]+"_CX_exportData");
                //克隆导出sheet
                XSSFSheet cloneExportSheet=workbook.cloneSheet(workbook.getSheetIndex(exportSheet));
                name=cloneExportSheet.getSheetName();
                //修改名字，名字必须有一定规则（导出的sheet的名字+"Data"=数据的sheet的名字）,这样VBA代码才能够将图表指向对应的data所在的sheet
                workbook.setSheetName(workbook.getSheetIndex(name),pointNames[i]+"_CX_export");
                //从模板读取模板的样式、内容，并填充到导出sheet的表数据部分
                //这里以一个单元格为例子
                XSSFCell testCell=cloneExportSheet.createRow(0).createCell(0);
                testCell.setCellStyle(templateSheet.getRow(0).getCell(0).getCellStyle());
                testCell.setCellValue(templateSheet.getRow(0).getCell(0).getStringCellValue());
                //往克隆的dataSheet里面插入数据,这样折线图会自动根据数据更新
                //创建8行数据
                cloneDataSheet.createRow(0);
                cloneDataSheet.createRow(1);
                cloneDataSheet.createRow(2);
                cloneDataSheet.createRow(3);
                cloneDataSheet.createRow(4);
                cloneDataSheet.createRow(5);
                cloneDataSheet.createRow(6);
                cloneDataSheet.createRow(7);
                //创建5列
                for (int j=0;j<8;j++){
                    if(j==0){
                        XSSFCell cell=cloneDataSheet.getRow(j).createCell(0);
                        cell.setCellValue("");
                        XSSFCell cell1=cloneDataSheet.getRow(j).createCell(1);
                        cell1.setCellValue("一月");
                        XSSFCell cell2=cloneDataSheet.getRow(j).createCell(2);
                        cell2.setCellValue("二月");
                        XSSFCell cell3=cloneDataSheet.getRow(j).createCell(3);
                        cell3.setCellValue("三月");
                        XSSFCell cell4=cloneDataSheet.getRow(j).createCell(4);
                        cell4.setCellValue("四月");
                    }else{
                        XSSFCell cell=cloneDataSheet.getRow(j).createCell(0);
                        cell.setCellValue(pointNames[i]+j);
                        XSSFCell cell1=cloneDataSheet.getRow(j).createCell(1);
                        cell1.setCellValue(Math.floor(Math.random()*10));
                        XSSFCell cell2=cloneDataSheet.getRow(j).createCell(2);
                        cell2.setCellValue(Math.floor(Math.random()*10));
                        XSSFCell cell3=cloneDataSheet.getRow(j).createCell(3);
                        cell3.setCellValue(Math.floor(Math.random()*10));
                        XSSFCell cell4=cloneDataSheet.getRow(j).createCell(4);
                        cell4.setCellValue(Math.floor(Math.random()*10));
                    }

                }

            }*/
            //经过验证，通过cloneSheet方法克隆的sheet是没有办法克隆sheet里面的VBA代码的，
            // 如果sheet里面有图的话，克隆之后用excel打开，excel会报错，并删除图的一些内容，导致图显示不出来
            // 所以不能用这种克隆sheet的方式来动态生成sheet
            //解决方法之一：如果要导出多个相同类型的sheet的话就改为导出多个excel文件
            //另外的解决办法：可以使用VBA来复制sheet，sheet的图和VBA代码貌似都可以复制过来，但是处理过程会比较复杂
            //====================================end===========================================


            XSSFSheet templateSheet=null;
            XSSFSheet dataSheet=null;
            XSSFSheet exportSheet=null;
            Iterator<Sheet> sheetIterator=workbook.sheetIterator();
            while(sheetIterator.hasNext()){
                XSSFSheet tempSheet=(XSSFSheet) sheetIterator.next();
                if(tempSheet.getSheetName().equals("表格模板")){
                    templateSheet=tempSheet;
                }
                if(tempSheet.getSheetName().equals("图的数据源")){
                    dataSheet=tempSheet;
                }
                if(tempSheet.getSheetName().equals("导出的sheet")){
                    exportSheet=tempSheet;
                }
            }
            System.out.println("a");
            //从模板读取模板的样式、内容，并填充到导出sheet的表数据部分
            //这里以一个单元格为例子
            XSSFCell testCell=exportSheet.createRow(0).createCell(0);
            testCell.setCellStyle(templateSheet.getRow(0).getCell(0).getCellStyle());
            testCell.setCellValue(templateSheet.getRow(0).getCell(0).getStringCellValue());
            //往excel的数据源的sheet里面插入数据,这样折线图会自动根据数据更新(在sheet中已经使用VBA写了处理逻辑，如果sheet被activate，那么就会自动根据数据的范围更新折线图)
            //创建8行数据
            dataSheet.createRow(0);
            dataSheet.createRow(1);
            dataSheet.createRow(2);
            dataSheet.createRow(3);
            dataSheet.createRow(4);
            dataSheet.createRow(5);
            dataSheet.createRow(6);
            dataSheet.createRow(7);
            //创建5列
            for (int j=0;j<8;j++){
                if(j==0){
                    XSSFCell cell=dataSheet.getRow(j).createCell(0);
                    cell.setCellValue("");
                    XSSFCell cell1=dataSheet.getRow(j).createCell(1);
                    cell1.setCellValue("一月");
                    XSSFCell cell2=dataSheet.getRow(j).createCell(2);
                    cell2.setCellValue("二月");
                    XSSFCell cell3=dataSheet.getRow(j).createCell(3);
                    cell3.setCellValue("三月");
                    XSSFCell cell4=dataSheet.getRow(j).createCell(4);
                    cell4.setCellValue("四月");
                }else{
                    XSSFCell cell=dataSheet.getRow(j).createCell(0);
                    cell.setCellValue("点"+j);
                    XSSFCell cell1=dataSheet.getRow(j).createCell(1);
                    cell1.setCellValue(Math.floor(Math.random()*10));
                    XSSFCell cell2=dataSheet.getRow(j).createCell(2);
                    cell2.setCellValue(Math.floor(Math.random()*10));
                    XSSFCell cell3=dataSheet.getRow(j).createCell(3);
                    cell3.setCellValue(Math.floor(Math.random()*10));
                    XSSFCell cell4=dataSheet.getRow(j).createCell(4);
                    cell4.setCellValue(Math.floor(Math.random()*10));
                }

            }

            //输出excel
            FileOutputStream fos = new FileOutputStream(new File("D:/exportExcel.xlsm"));
            workbook.write(fos);
            fos.close();
            workbook.close();
            System.out.println("");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}