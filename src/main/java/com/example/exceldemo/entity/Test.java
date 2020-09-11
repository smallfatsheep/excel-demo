package com.example.exceldemo.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;

import java.util.Date;

/**
 * @Author: huangwc
 * @Description:
 * @Date: 2020/09/07 16:56:14
 * @Version: 1.0
 */
@Setter
@Getter
@Data
public class Test {

    @Excel(name = "表名")
    private String name;
    @Excel(name = "工程名")
    private String pjname;
    @Excel(name = "项目")
    private String project;
    @Excel(name = "机器")
    private String machine;
    @Excel(name = "时间")
    private Date date;

}
