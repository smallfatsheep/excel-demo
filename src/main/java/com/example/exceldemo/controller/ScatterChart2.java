package com.example.exceldemo.controller;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Random;

import static org.openxmlformats.schemas.drawingml.x2006.chart.STOrientation.MIN_MAX;
import static org.openxmlformats.schemas.drawingml.x2006.chart.STTickLblPos.NEXT_TO;

/**
 * @Author: huangwc
 * @Description:
 * @Date: 2020/09/03 09:44:27
 * @Version: 1.0
 */
public class ScatterChart2 {
    public static void main(String[] args) throws IOException {
        try {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.createSheet("sheet1");

            Row row;
            Cell cell;
            for (int r = 0; r < 105; r++) {
                row = sheet.createRow(r);
                cell = row.createCell(0);
                cell.setCellValue("S" + r);
                cell = row.createCell(1);
                cell.setCellValue(new Random().nextInt(100));
            }

            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 5, 24, 45);

            XSSFChart chart = drawing.createChart(anchor);

            chart.setTitleText("预选赛项目得分分布图");
            chart.setAutoTitleDeleted(false);

            CTChart ctChart = chart.getCTChart();
            ctChart.addNewPlotVisOnly().setVal(true);
            ctChart.addNewDispBlanksAs().setVal(STDispBlanksAs.Enum.forInt(2));
            ctChart.addNewShowDLblsOverMax().setVal(false);

            // 创建一个散点图
            CTPlotArea ctPlotArea = ctChart.getPlotArea();

            CTScatterChart scatterChart = ctPlotArea.addNewScatterChart();
            scatterChart.addNewScatterStyle().setVal(STScatterStyle.LINE_MARKER);
            scatterChart.addNewVaryColors().setVal(false); // 不允许自定义颜色
            scatterChart.addNewAxId().setVal(123456);
            scatterChart.addNewAxId().setVal(123457);

            CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
            ctCatAx.addNewAxId().setVal(123456);
            CTScaling ctScaling = ctCatAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(MIN_MAX);
            ctCatAx.addNewDelete().setVal(false);
            //ctCatAx.addNewAxPos().setVal(STAxPos.B);
            ctCatAx.addNewCrossAx().setVal(123457);
            ctCatAx.addNewTickLblPos().setVal(NEXT_TO);

            // 设置Y坐标
            CTValAx ctValAx = ctPlotArea.addNewValAx();
            ctValAx.addNewAxId().setVal(123457);
            CTScaling ctScaling1 = ctValAx.addNewScaling();
            ctScaling1.addNewOrientation().setVal(MIN_MAX);
            ctValAx.addNewDelete().setVal(false);
            //ctValAx.addNewAxPos().setVal(STAxPos.B);
            ctValAx.addNewCrossAx().setVal(123456);
            // Y轴的对比线
            CTShapeProperties ctShapeProperties = ctValAx.addNewMajorGridlines().addNewSpPr();
            CTLineProperties ctLineProperties = ctShapeProperties.addNewLn();
            ctLineProperties.setW(9525);
            ctLineProperties.setCap(STLineCap.Enum.forInt(3));
            ctLineProperties.setCmpd(STCompoundLine.Enum.forInt(1));
            ctLineProperties.setAlgn(STPenAlignment.Enum.forInt(1));
            // 不显示Y轴上的坐标刻度线
            //ctValAx.addNewMajorTickMark().setVal(STTickMark.NONE);
            //ctValAx.addNewMinorTickMark().setVal(STTickMark.NONE);
            //ctValAx.addNewTickLblPos().setVal(NEXT_TO);

            // 设置散点图内的信息
            CTScatterSer ctScatterSer = scatterChart.addNewSer();
            ctScatterSer.addNewIdx().setVal(0);
            ctScatterSer.addNewOrder().setVal(0);
            // 去掉连接线
            ctPlotArea.getScatterChartArray(0).getSerArray(0).addNewSpPr().addNewLn().addNewNoFill();

            // 设置散点图各图例的显示
            CTDLbls ctdLbls = scatterChart.addNewDLbls();
            ctdLbls.addNewShowVal().setVal(true);
            ctdLbls.addNewShowLegendKey().setVal(false);
            ctdLbls.addNewShowSerName().setVal(false);
            ctdLbls.addNewShowCatName().setVal(false);
            ctdLbls.addNewShowPercent().setVal(false);
            ctdLbls.addNewShowBubbleSize().setVal(false);
            // 设置标记的样式
            CTMarker ctMarker = ctScatterSer.addNewMarker();
            ctMarker.addNewSymbol().setVal(STMarkerStyle.Enum.forInt(3));
            ctMarker.addNewSize().setVal((short) 5);
            CTShapeProperties ctShapeProperties1 = ctMarker.addNewSpPr();
            ctShapeProperties1.addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.Enum.forInt(5));
            CTLineProperties ctLineProperties1 = ctShapeProperties1.addNewLn();
            ctLineProperties1.setW(9525);
            ctLineProperties1.addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.Enum.forInt(5));

            CTAxDataSource ctAxDataSource = ctScatterSer.addNewXVal();
            CTStrRef ctStrRef = ctAxDataSource.addNewStrRef();
            ctStrRef.setF("sheet1!$A$1:$A$100");
            CTNumDataSource ctNumDataSource = ctScatterSer.addNewYVal();
            CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
            ctNumRef.setF("sheet1!$B$1:$B$100");
            FileOutputStream fileOut = new FileOutputStream("D:\\out.xlsx");
            wb.write(fileOut);
            fileOut.close();
        }catch (Exception e){
            e.printStackTrace();
        }




    }
}
