package com.example.demoio.core.bean;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

@Data
public class Person {


    private String id;

    @Excel(name = "姓名",width= 20)
    private String userName;

    @Excel(name = "密码",width= 20)
    private String passwd;

    @Excel(name = "性别",width = 20, replace = {"男_1","女_0"})
    private String age;

    @Excel(name = "电话",width= 20)
    private String phone;


    public Person(String id, String userName, String passwd, String age, String phone) {
        this.id = id;
        this.userName = userName;
        this.passwd = passwd;
        this.age = age;
        this.phone = phone;
    }

    public Person() {
    }

    /**
     * 模拟数据
     * @return
     */
    public static Person setPerson() {
        Person person = new Person();
        person.setId(String.valueOf(System.currentTimeMillis()));
        person.setAge("21");
        person.setPasswd("123456");
        person.setUserName("lisi");
        person.setPhone("12312341234");
        return person;
    }
}
