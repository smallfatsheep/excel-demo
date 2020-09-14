package com.example.exceldemo.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;

/**
 * @Author: huangwc
 * @Description:
 * @Date: 2020/09/07 17:02:09
 * @Version: 1.0
 */
@Setter
@Getter
@Data
public class Datas {
    //@Excel(name = "深度  （m）" ,needMerge = true)
    private double deep;
    //@Excel(name = "本次变化（mm）",needMerge = true)
    private double change;
    //@Excel(name = "本次速率(mm/d)",needMerge = true)
    private double speed;
    //@Excel(name = "累计变化(mm)",needMerge = true)
    private double addupchange;
    private String test;

}
