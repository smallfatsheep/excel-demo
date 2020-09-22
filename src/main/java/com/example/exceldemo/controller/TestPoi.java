package com.example.exceldemo.controller;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScatterChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.main.STSchemeColorVal;

import java.io.*;

/**
 * @Author: huangwc
 * @Description:
 * @Date: 2020/09/14 10:40:16
 * @Version: 1.0
 */
public class TestPoi {
    public static void main(String[] args) {
        try (XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(new File("D:/home/excel/poi.xlsx")));)
        {
            XSSFSheet sheet = wb.getSheet("Sheet1");
            CellStyle stringSeptailStyle = wb.createCellStyle();
            stringSeptailStyle.setDataFormat(wb.createDataFormat().getFormat("0.00"));

            //获得总行数
            int rowNum=sheet.getPhysicalNumberOfRows();
            //获得总列数
            int coloumNum=sheet.getRow(0).getPhysicalNumberOfCells();
            System.out.println("获得总行数 :" + rowNum);
            System.out.println("获得总列数 :" + coloumNum);

            //创建画布
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            //8个参数分别代表左上角和右下角所在单元格的坐标和偏移量
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 5, 4 ,10, 13);
            //创建图表
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("J5变化曲线图");
            //chart.setAutoTitleDeleted(false);
            chart.setTitleOverlay(false);

            //获取图列
            XDDFChartLegend legend = chart.getOrAddLegend();
            //图列位置
            legend.setPosition(LegendPosition.RIGHT);

            //创建轴
            XDDFValueAxis bottomAxis = chart.createValueAxis(AxisPosition.LEFT);
            //设置轴的数值间隔、最大值、最小值
            bottomAxis.setMinorUnit(0.5d);
            bottomAxis.setMajorUnit(0.5d);
            bottomAxis.setMinimum(0.0d);
            bottomAxis.setMaximum(10.0d);
            //轴标题
            bottomAxis.setTitle("孔深(m)");
            //轴方向
            bottomAxis.setOrientation(AxisOrientation.MAX_MIN);
            //一条轴对齐另一条轴的最大或最小值
            bottomAxis.setCrosses(AxisCrosses.MIN);
            //标签位置
            //bottomAxis.setTickLabelPosition(AxisTickLabelPosition.NEXT_TO);
            bottomAxis.setMajorTickMark(AxisTickMark.NONE);

            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.BOTTOM);
            leftAxis.setMaximum(10.0d);
            leftAxis.setMinimum(-4.0d);
            leftAxis.setMinorUnit(2.0d);
            leftAxis.setMajorUnit(2.0d);
            leftAxis.setNumberFormat("d");
            leftAxis.setTitle("位移(m)");
            leftAxis.setOrientation(AxisOrientation.MIN_MAX);
            leftAxis.setCrosses(AxisCrosses.MIN);
            leftAxis.setMajorTickMark(AxisTickMark.NONE);

            //对应三行数据
            XDDFDataSource <Double> xs1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,new CellRangeAddress(4, 13, 3, 3));
            XDDFDataSource <Double> xs2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet,new CellRangeAddress(4, 13, 1, 1));
            XDDFNumericalDataSource <Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet,new CellRangeAddress(4, 13, 0, 0));

            XDDFScatterChartData data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, leftAxis, bottomAxis);
            data.setStyle(ScatterStyle.LINE);
            data.setVaryColors(false);

            //增加系列
            XDDFScatterChartData.Series series1 = (XDDFScatterChartData.Series) data.addSeries(xs1, ys);
            series1.getShapeProperties();
            series1.setMarkerStyle(MarkerStyle.STAR);
            series1.setTitle("2020/5/24", null);
            XDDFScatterChartData.Series series2 = (XDDFScatterChartData.Series) data.addSeries(xs2, ys);
            series2.setMarkerStyle(MarkerStyle.SQUARE);
            series2.setTitle("2020/8/15", null);
            series2.setMarkerSize((short) 6);
            chart.plot(data);

            CTChart ctChart = chart.getCTChart();
            CTPlotArea ctPlotArea = ctChart.getPlotArea();
            CTScatterChart ctScatterChart = ctPlotArea.getScatterChartArray(0);

            CTValAx ctValAx1 = ctPlotArea.getValAxArray(0);
            CTValAx ctValAx2 = ctPlotArea.getValAxArray(1);
            ctValAx1.addNewMajorGridlines().addNewSpPr().addNewLn().addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.Enum.forInt(3));// 显示主要网格线，并设置颜色
            ctValAx2.addNewMajorGridlines().addNewSpPr().addNewLn().addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.Enum.forInt(3));// 显示主要网格线，并设置颜色

            solidLineSeries(data, 0, PresetColor.AQUA);
            solidLineSeries(data, 1, PresetColor.TURQUOISE);

            try (FileOutputStream fileOut = new FileOutputStream("D:/home/excel/555.xlsx")) {
                wb.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void solidLineSeries(XDDFChartData data, int index, PresetColor color) {
        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(color));
        XDDFLineProperties line = new XDDFLineProperties();
        line.setFillProperties(fill);
        XDDFChartData.Series series = data.getSeries().get(index);
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setLineProperties(line);
        series.setShapeProperties(properties);
    }

}

