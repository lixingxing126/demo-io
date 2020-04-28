package com.example.demoio.execl;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import com.example.demoio.core.bean.Person;
import com.example.demoio.core.util.ExcelUtil;
import io.swagger.annotations.Api;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;


/**
 * 使用easypoi导出execl
 * 适合导出简单的execl文件
 * @author lxx
 * @date 2020年4月24日
 */
@Slf4j
@Api(value = "/easypoi", tags = "execl - easypoi")
@RestController
@RequestMapping("/easypoi")
public class EasyPOIController {

    /**
     * 导出文件流
     */
    @GetMapping("/exportIO")
    public void exportIO(HttpServletResponse res) {
        try {
            List<Person> collect = new ArrayList<>();
            collect.add(Person.setPerson());
            collect.add(Person.setPerson());
            collect.add(Person.setPerson());
            Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("导出execl", "sheet"),
                    Person.class, collect);
            res.setContentType("application/octet-stream");
            //设置Excel文件名
            res.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode("案件导出.xls", "UTF-8"));
            try {
                res.flushBuffer();
                workbook.write(res.getOutputStream());
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 导出到本地目录下
     * @param resp
     */
    @GetMapping("/exportLocal")
    public void exportLocal(HttpServletResponse resp) {
        try {
            String path = "C:\\Users\\Administrator\\Desktop\\person.xlsx";
            List<Person> collect = new ArrayList<>();
            collect.add(Person.setPerson());
            collect.add(Person.setPerson());
            collect.add(Person.setPerson());
            Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("导出execl", "sheet"),
                    Person.class, collect);
            resp.setContentType("application/octet-stream");
            //设置Excel文件名
            resp.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode("案件导出.xls", "UTF-8"));
            try {
                resp.flushBuffer();
                FileOutputStream fos = new FileOutputStream(path);
                workbook.write(fos);
                fos.flush();
                fos.close();
                log.info("execl导出到" + path + "目录下。");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 导入execl
     * @param multipartFile
     * @return
     */
    @GetMapping("/importExecl")
    public Boolean importExecl(MultipartFile multipartFile) {
        try {
            List<Person> excelList = ExcelUtil.importExcel(multipartFile, 3, 3, Person.class);
            for (Person person : excelList) {
                log.info(person.toString());
            }
        } catch (Exception e) {
            throw new RuntimeException("导入失败");
        }
        return Boolean.TRUE;
    }
}
