package com.example.demoio.execl;

import com.example.demoio.core.bean.Person;
import io.swagger.annotations.Api;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

/**
 * 使用easypoi导出execl
 * 适合导出简单的execl文件
 * @author lxx
 * @date 2020年4月24日
 */
@Slf4j
@Api(value = "/poi/import", tags = "导出execl - poi")
@RestController
@RequestMapping("/poi/import")
public class POIController {

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";

    /**
     * 导入
     * @param file
     * @return
     */
    @GetMapping("/import")
    public boolean importExcel(MultipartFile file) {
        if (!XLS.equals(file.getOriginalFilename().split("\\.")[1]) && !XLSX.equals(file.getOriginalFilename().split("\\.")[1])) {
            log.info("请导入execl格式的文件");
        }
        try {
            //根据指定的文件输入流导入Excel从而产生Workbook对象
            Workbook wb0 = new XSSFWorkbook(file.getInputStream());
            //获取Excel文档中的第一个表单
            Sheet sht0 = wb0.getSheetAt(0);
            //对Sheet中的每一行进行迭代
            for (Row r : sht0) {
                //跳过标题列
                //如果当前行的行号（从0开始）未达到2（第三行）则从新循环
//                if (r.getRowNum() < 2) {
//                    continue;
//                }
                //创建实体类
                Person person = new Person();
                //取出当前行第1个单元格数据，并封装在info实体stuName属性上
                person.setAge(String.valueOf(r.getCell(1)));
                person.setId(String.valueOf(r.getCell(2)));
                person.setPhone(String.valueOf(r.getCell(3)));
                person.setPasswd(String.valueOf(r.getCell(4)));
                person.setUserName(String.valueOf(r.getCell(5)));
                log.info(person.toString());
            }
            return true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;
    }

}
