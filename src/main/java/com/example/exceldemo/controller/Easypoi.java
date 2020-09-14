package com.example.exceldemo.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.example.exceldemo.entity.Datas;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * @Author: huangwc
 * @Description:
 * @Date: 2020/09/07 15:07:47
 * @Version: 1.0
 */
public class Easypoi {
    public void easypoi(String filename,Map<String, Object> map) throws IOException {
        TemplateExportParams params = new TemplateExportParams(
                filename);
        //params.setHeadingStartRow(3);
        //params.setHeadingRows(1);
        params.setSheetName(new String[]{"Sheet1"});
        Map<String, Object> map1 = new HashMap<String, Object>();
        map1.put("date", new Date());
        map1.put("name", "土体侧向");
        map1.put("pjname", "祈福新邨BC0104082地块基坑监测工程");
        map1.put("machine", "武汉基深测斜仪CX-3C");
        map1.put("times", 100);
        map1.put("project", "土体侧向位移C5");
        map1.put("tablename", "土体侧向位移监测成果表");
        map1.put("username1", "黄斌");
        map1.put("username2", "龙晋航");
        map1.put("username3", "李瑞华");
        map1.put("other", "1、“+”表示向基坑内位移，“-”表示向基坑外位移；\n"+
                "2、累计预警值为35mm，变化速率为6mm/d，控制值为44mm。");
        map1.put("end", "本期监测数据显示，各监测点数据变化均未超过设计预警值。");
        map1.put("status", "一三号楼正在进行板底加工，裙楼正在进行设计锚杆施工，二号楼正在进行锚杆施工。");
        List<Datas> list = new ArrayList<Datas>();

        for (int i = 0; i < 10; i++) {
            Datas entity = new Datas();
            entity.setDeep(i / 2 + 0.8);
            entity.setChange(i / 4 + 0.8);
            entity.setSpeed(i / 6 + 0.8);
            entity.setAddupchange(i / 8 + 0.8);
            entity.setTest("");
            list.add(entity);
        }

        map1.put("maplist", list);

        Workbook workbook = ExcelExportUtil.exportExcel(params, map1);
        File savefile = new File("D:/home/excel/");
        //workbook.getSheet("成果表").addMergedRegion(new CellRangeAddress(3, 13, 4, 10));
        //PoiMergeCellUtil.mergeCells(workbook.getSheet("成果表"), 4, 4);
        CellRangeAddress cellRangePlanNo = new CellRangeAddress(3, 13, 4, 10);
        //CellStyle cellStyle = workbook.createCellStyle();
        // 使用字符串定义格式
        //cellStyle.setDataFormat((short) BuiltinFormats.getBuiltinFormat("0.0"));
        // BuiltinFormats._formats 数组中的下标
        //cellStyle.setDataFormat((short) 1);
        workbook.getSheetAt(0).addMergedRegion(cellRangePlanNo);
        RegionUtil.setBorderBottom(BorderStyle.THIN, cellRangePlanNo, workbook.getSheetAt(0));
        RegionUtil.setBorderLeft(BorderStyle.THIN, cellRangePlanNo, workbook.getSheetAt(0));
        RegionUtil.setBorderRight(BorderStyle.THIN, cellRangePlanNo, workbook.getSheetAt(0));
        RegionUtil.setBorderTop(BorderStyle.THIN, cellRangePlanNo, workbook.getSheetAt(0));
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("D:/home/excel/result.xlsx");
        workbook.write(fos);
        fos.close();
    }
}
