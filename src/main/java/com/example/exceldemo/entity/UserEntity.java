package com.example.exceldemo.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

import java.io.Serializable;

/**
 * @Author: huangwc
 * @Description:
 * @Date: 2020/09/16 17:47:51
 * @Version: 1.0
 */
@Data

public class UserEntity implements Serializable {

    private static final long serialVersionUID = 1L;
    private String idx;

    @Excel(name = "籍贯", mergeVertical = true, width = 50)
    private String nativeStr;
    @Excel(name = "姓名", width = 20)
    private String userName;
    @Excel(name = "年龄")
    private int age;
    @Excel(name = "地址", width = 50)
    private String addr;
}
