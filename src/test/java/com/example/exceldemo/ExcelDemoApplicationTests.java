package com.example.exceldemo;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelStyleType;
import cn.afterturn.easypoi.util.PoiMergeCellUtil;
import com.example.exceldemo.entity.Datas;
import com.example.exceldemo.entity.UserEntity;
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
                "D://home/excel/test.xlsx");
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

//        for (int i = 0; i < 10; i++) {
//            Datas entity = new Datas();
//            entity.setDeep(i/2 + 0.8);
//            entity.setChange(i/4 + 0.8);
//            entity.setSpeed(i/6 + 0.8);
//            entity.setAddupchange(i/8 + 0.8);
//            entity.setTest("");
//            list.add(entity);
//        }

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
        /*在创建TemplateExportParams模版对象时，若需要输出多sheet的话，
        需要在指定模版路径后，追加一个参数，
        默认值为false，设置为true即表示会输出模版中的全部sheet，否则只会输出第一个sheet*/
        FileOutputStream fos = new FileOutputStream("D:/home/excel/555.xlsx");
        workbook.write(fos);
        fos.close();
    }


    @Test
    public void one() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "D://home/excel/for_Col.xlsx",2,1);
        params.setColForEach(true);
        params.setSheetName(new String[]{"sheet_1","sheet_2"});
        Map<String, Object> value = new HashMap<String, Object>();
        List<Map<String, Object>> colList = new ArrayList<Map<String, Object>>();
        List<Map<String, Object>> colList1 = new ArrayList<Map<String, Object>>();

        //先处理表头
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("name", "第125次");
        map.put("time", new Date());
        map.put("zq", "深");
        map.put("cw", "宽");
        map.put("tj", "长");
        map.put("zqmk", "t.zq_xm");
        map.put("cwmk", "t.cw_xm");
        map.put("tjmk", "t.tj_xm");
        colList.add(map);

        map = new HashMap<String, Object>();
        map.put("name", "第126次");
        map.put("time", new Date());
        map.put("zq", "深");
        map.put("cw", "宽");
        map.put("tj", "长");
        map.put("zqmk", "t.zq_xh");
        map.put("cwmk", "t.cw_xh");
        map.put("tjmk", "t.tj_xh");
        colList.add(map);

        Map<String, Object> map1 = new HashMap<String, Object>();
        map1.put("name", "第125次");
        map1.put("time", new Date());
        map1.put("zq", "深");
        map1.put("cw", "宽");
        map1.put("tj", "长");
        map1.put("zqmk", "t.zq_xm");
        map1.put("cwmk", "t.cw_xm");
        map1.put("tjmk", "t.tj_xm");
        colList1.add(map1);


        value.put("colList", colList);
        value.put("colList2", colList);
        value.put("colList3", colList);
        List<Map<String, Object>> valList = new ArrayList<Map<String, Object>>();
        for (int i = 0; i < 10; i++){
            map = new HashMap<String, Object>();
            map.put("one", "基坑顶部");
            map.put("two", "PO"+(i+1));
            map.put("zq_xm","zqxm"+i);
            map.put("cw_xm","cwxm"+i);
            map.put("tj_xm","tjxm"+i);
            map.put("zq_xh","zqxh"+i);
            map.put("cw_xh","cwxh"+i);
            map.put("tj_xh","tjxh"+i);
            valList.add(map);
        }


        value.put("name", "基坑顶部");
        value.put("valList", valList);
        value.put("valList2", valList);
        value.put("valList3", valList);
        Workbook book = ExcelExportUtil.exportExcel(params, value);
        PoiMergeCellUtil.mergeCells(book.getSheetAt(0), 1, 0);
        PoiMergeCellUtil.mergeCells(book.getSheetAt(1), 1, 0);
        FileOutputStream fos = new FileOutputStream("D://home/excel/result2.xlsx");
        book.write(fos);
        fos.close();
    }
    @Test
    public void test2() throws Exception {

        Map<String, Object> map = new HashMap<String, Object>();
        List<UserEntity> list = new ArrayList<>();

        for (int i = 0; i < 10; i++) {
            UserEntity entity = new UserEntity();
            entity.setIdx(i + "");
            entity.setNativeStr("广东梅州");
            entity.setUserName("Mrs Ling_" + i);
            if (i > 4) {
                entity.setUserName("Mrs Ling");
                entity.setIdx("5");
            }
            entity.setAge(16 + i);
            entity.setAddr("广东梅州_" + i);
            list.add(entity);
        }

        for (int i = 0; i < 10; i++) {
            UserEntity entity = new UserEntity();
            entity.setIdx(i + "");
            entity.setNativeStr("广西玉林");
            entity.setUserName("Mr Feng_" + i);
            if (i > 4) {
                entity.setUserName("Mr Feng");
                entity.setIdx("5");
            }
            entity.setAge(21 + i);
            entity.setAddr("广西玉林_" + i);
            list.add(entity);
        }

        map.put("entityList", list);

        TemplateExportParams params = new TemplateExportParams(
                "D://home/excel/test3.xlsx");
        ExcelExportUtil.exportExcel(params, map);
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);

        PoiMergeCellUtil.mergeCells(workbook.getSheetAt(0), 1, 0, 1,3,2,4);
        File saveFolder = new File("D:/home/excel/");
        if (!saveFolder.exists()) {
            saveFolder.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("D://home/excel/fengling_test_export" + System.currentTimeMillis() + ".xlsx");
        workbook.write(fos);
        fos.close();
    }
}
