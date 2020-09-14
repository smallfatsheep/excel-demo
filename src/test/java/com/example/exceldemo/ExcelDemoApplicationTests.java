package com.example.exceldemo;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelStyleType;
import cn.afterturn.easypoi.util.PoiMergeCellUtil;
import com.example.exceldemo.entity.Datas;
import com.google.common.collect.Lists;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import static org.openxmlformats.schemas.drawingml.x2006.chart.STOrientation.MIN_MAX;
import static org.openxmlformats.schemas.drawingml.x2006.chart.STTickLblPos.NEXT_TO;

@SpringBootTest
class ExcelDemoApplicationTests {

    @Test
    void contextLoads() {
    }
    @Test
    public void test() throws IOException {
        TemplateExportParams params = new TemplateExportParams(
                "D://home/excel/test.xls");
        params.setSheetName(new String[]{"祈福"});
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("date", new Date());
        map.put("tablename","土体侧向位移监测成果表");
        map.put("name", "土体侧向");
        map.put("pjname", "祈福工程");
        map.put("machine", "武汉基深");
        map.put("times", 100);
        map.put("project", "土体侧向位移");
        List<Map<String, Integer>> listMap = new ArrayList<Map<String, Integer>>();
        for (int i = 0; i < 4; i++) {
            Map<String, Integer> lm = new HashMap<String, Integer>();
            lm.put("deep", i + 1);
            lm.put("change", i * 10000);
            lm.put("addupchange", 12);
            lm.put("speed", 21);
            listMap.add(lm);
        }
        map.put("maplist", listMap);

        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        File savefile = new File("D:/home/excel/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("D:/home/excel/222.xls");
        workbook.write(fos);
        fos.close();

    }
    @Test
    public void test1() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "D://home/excel/test.xlsx");
        params.setHeadingStartRow(3);
        params.setHeadingRows(1);
        params.setSheetName(new String[]{"成果表"});
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("date", new Date());
        map.put("name", "土体侧向");
        map.put("pjname", "祈福工程");
        map.put("machine", "武汉基深");
        map.put("times", 100);
        map.put("project", "土体侧向位移");
        map.put("tablename", "土体侧向位移监测成果表");
        map.put("username1", "黄斌");
        map.put("username2", "黄斌");
        map.put("username3", "黄斌");
        map.put("other", "大大方方烦烦烦不方便很讨dfadfdfadsf");
        map.put("end", "dfadfasd大噶啊啊v觉哦i给你发了白马非马不来了");
        map.put("status", "打飞机拉萨大家辣椒辣女哦人能够让老师的");
        List<Datas> list = new ArrayList<Datas>();

        for (int i = 0; i < 10; i++) {
            Datas entity = new Datas();
            entity.setDeep(i/2 + 0.8);
            entity.setChange(i/4 + 0.8);
            entity.setSpeed(i/6 + 0.8);
            entity.setAddupchange(i/8 + 0.8);
            entity.setTest("");
            list.add(entity);
        }

        map.put("maplist", list);

        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        File savefile = new File("D:/home/excel/");
        //workbook.getSheet("成果表").addMergedRegion(new CellRangeAddress(3, 13, 4, 10));
        //PoiMergeCellUtil.mergeCells(workbook.getSheet("成果表"), 4, 4);
        CellRangeAddress cellRangePlanNo = new CellRangeAddress(3, 13, 4, 10);
        workbook.getSheet("成果表").addMergedRegion(cellRangePlanNo);
        RegionUtil.setBorderBottom(BorderStyle.THIN, cellRangePlanNo, workbook.getSheet("成果表"));
        RegionUtil.setBorderLeft(BorderStyle.THIN, cellRangePlanNo, workbook.getSheet("成果表"));
        RegionUtil.setBorderRight(BorderStyle.THIN, cellRangePlanNo, workbook.getSheet("成果表"));
        RegionUtil.setBorderTop(BorderStyle.THIN, cellRangePlanNo, workbook.getSheet("成果表"));
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("D:/home/excel/555.xlsx");
        workbook.write(fos);
        fos.close();
    }
    @Test
    public void createScatterChart() throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("散点图");

        Row row;
        Cell cell;
        for (int r = 0; r < 105; r++) {
            row = sheet.createRow(r);
            cell = row.createCell(0);
            cell.setCellValue("S" + r);
            cell = row.createCell(1);
            cell.setCellValue(100);
        }

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 0, 21, 40);

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
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(123457);
        ctCatAx.addNewTickLblPos().setVal(NEXT_TO);

        // 设置Y坐标
        CTValAx ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(123457);
        CTScaling ctScaling1 = ctValAx.addNewScaling();
        ctScaling1.addNewOrientation().setVal(MIN_MAX);
        ctValAx.addNewDelete().setVal(false);
        ctValAx.addNewAxPos().setVal(STAxPos.B);
        ctValAx.addNewCrossAx().setVal(123456);
        // Y轴的对比线
        CTShapeProperties ctShapeProperties = ctValAx.addNewMajorGridlines().addNewSpPr();
        CTLineProperties ctLineProperties = ctShapeProperties.addNewLn();
        ctLineProperties.setW(9525);
        ctLineProperties.setCap(STLineCap.Enum.forInt(3));
        ctLineProperties.setCmpd(STCompoundLine.Enum.forInt(1));
        ctLineProperties.setAlgn(STPenAlignment.Enum.forInt(1));
        // 不显示Y轴上的坐标刻度线
        ctValAx.addNewMajorTickMark().setVal(STTickMark.NONE);
        ctValAx.addNewMinorTickMark().setVal(STTickMark.NONE);
        ctValAx.addNewTickLblPos().setVal(NEXT_TO);

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
        ctStrRef.setF("散点图!$A$1:$A$100");
        CTNumDataSource ctNumDataSource = ctScatterSer.addNewYVal();
        CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
        ctNumRef.setF("散点图!$B$1:$B$100");

        System.out.println(ctChart);

        FileOutputStream fileOut = new FileOutputStream("D:\\out.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }


//    @Test
//    public void createScatterChart() throws IOException {
//        XSSFWorkbook wb = new XSSFWorkbook();
//        XSSFSheet sheet = wb.createSheet("散点图");
//
//        Row row;
//        Cell cell;
//        for (int r = 0; r < 105; r++) {
//            row = sheet.createRow(r);
//            cell = row.createCell(0);
//            cell.setCellValue("S" + r);
//            cell = row.createCell(1);
//            cell.setCellValue(RandomUtils.nextInt(1,100));
//        }
//
//        XSSFDrawing drawing = sheet.createDrawingPatriarch();
//        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 0, 21, 40);
//
//        XSSFChart chart = drawing.createChart(anchor);
//
//        chart.setTitleText("预选赛项目得分分布图");
//        chart.setAutoTitleDeleted(false);
//
//        CTChart ctChart = chart.getCTChart();
//        ctChart.addNewPlotVisOnly().setVal(true);
//        ctChart.addNewDispBlanksAs().setVal(STDispBlanksAs.Enum.forInt(2));
//        ctChart.addNewShowDLblsOverMax().setVal(false);
//
//        // 创建一个散点图
//        CTPlotArea ctPlotArea = ctChart.getPlotArea();
//
//        CTScatterChart scatterChart = ctPlotArea.addNewScatterChart();
//        scatterChart.addNewScatterStyle().setVal(STScatterStyle.LINE_MARKER);
//        scatterChart.addNewVaryColors().setVal(false); // 不允许自定义颜色
//        scatterChart.addNewAxId().setVal(123456);
//        scatterChart.addNewAxId().setVal(123457);
//
//        CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
//        ctCatAx.addNewAxId().setVal(123456);
//        CTScaling ctScaling = ctCatAx.addNewScaling();
//        ctScaling.addNewOrientation().setVal(MIN_MAX);
//        ctCatAx.addNewDelete().setVal(false);
//        ctCatAx.addNewAxPos().setVal(STAxPos.B);
//        ctCatAx.addNewCrossAx().setVal(123457);
//        ctCatAx.addNewTickLblPos().setVal(NEXT_TO);
//
//        // 设置Y坐标
//        CTValAx ctValAx = ctPlotArea.addNewValAx();
//        ctValAx.addNewAxId().setVal(123457);
//        CTScaling ctScaling1 = ctValAx.addNewScaling();
//        ctScaling1.addNewOrientation().setVal(MIN_MAX);
//        ctValAx.addNewDelete().setVal(false);
//        ctValAx.addNewAxPos().setVal(STAxPos.B);
//        ctValAx.addNewCrossAx().setVal(123456);
//        // Y轴的对比线
//        CTShapeProperties ctShapeProperties = ctValAx.addNewMajorGridlines().addNewSpPr();
//        CTLineProperties ctLineProperties = ctShapeProperties.addNewLn();
//        ctLineProperties.setW(9525);
//        ctLineProperties.setCap(STLineCap.Enum.forInt(3));
//        ctLineProperties.setCmpd(STCompoundLine.Enum.forInt(1));
//        ctLineProperties.setAlgn(STPenAlignment.Enum.forInt(1));
//        // 不显示Y轴上的坐标刻度线
//        ctValAx.addNewMajorTickMark().setVal(STTickMark.NONE);
//        ctValAx.addNewMinorTickMark().setVal(STTickMark.NONE);
//        ctValAx.addNewTickLblPos().setVal(NEXT_TO);
//
//        // 设置散点图内的信息
//        CTScatterSer ctScatterSer = scatterChart.addNewSer();
//        ctScatterSer.addNewIdx().setVal(0);
//        ctScatterSer.addNewOrder().setVal(0);
//        // 去掉连接线
//        ctPlotArea.getScatterChartArray(0).getSerArray(0).addNewSpPr().addNewLn().addNewNoFill();
//
//        // 设置散点图各图例的显示
//        CTDLbls ctdLbls = scatterChart.addNewDLbls();
//        ctdLbls.addNewShowVal().setVal(true);
//        ctdLbls.addNewShowLegendKey().setVal(false);
//        ctdLbls.addNewShowSerName().setVal(false);
//        ctdLbls.addNewShowCatName().setVal(false);
//        ctdLbls.addNewShowPercent().setVal(false);
//        ctdLbls.addNewShowBubbleSize().setVal(false);
//        // 设置标记的样式
//        CTMarker ctMarker = ctScatterSer.addNewMarker();
//        ctMarker.addNewSymbol().setVal(STMarkerStyle.Enum.forInt(3));
//        ctMarker.addNewSize().setVal((short) 5);
//        CTShapeProperties ctShapeProperties1 = ctMarker.addNewSpPr();
//        ctShapeProperties1.addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.Enum.forInt(5));
//        CTLineProperties ctLineProperties1 = ctShapeProperties1.addNewLn();
//        ctLineProperties1.setW(9525);
//        ctLineProperties1.addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.Enum.forInt(5));
//
//        CTAxDataSource ctAxDataSource = ctScatterSer.addNewXVal();
//        CTStrRef ctStrRef = ctAxDataSource.addNewStrRef();
//        ctStrRef.setF("散点图!$A$1:$A$100");
//        CTNumDataSource ctNumDataSource = ctScatterSer.addNewYVal();
//        CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
//        ctNumRef.setF("散点图!$B$1:$B$100");
//
//        System.out.println(ctChart);
//
//        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\user\\Desktop\\out.xlsx");
//        wb.write(fileOut);
//        fileOut.close();
//    }
}
