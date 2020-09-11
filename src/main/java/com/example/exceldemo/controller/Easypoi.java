package com.example.exceldemo.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.example.exceldemo.entity.Datas;
import org.apache.poi.ss.usermodel.Workbook;

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
    public void easypoi() throws IOException {
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
            list.add(entity);
        }

        map.put("maplist", list);
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        File savefile = new File("D:/home/excel/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("D:/home/excel/555.xlsx");
        workbook.write(fos);
        fos.close();
    }
}
