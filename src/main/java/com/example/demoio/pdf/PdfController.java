package com.example.demoio.pdf;

import com.example.demoio.core.util.ExportPDF;
import com.google.common.collect.Lists;
import com.itextpdf.text.pdf.PdfReader;
import io.swagger.annotations.Api;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

@Slf4j
@Api(value = "/pdf", tags = "PDF - itextPDF")
@RestController
@RequestMapping("/pdf")
public class PdfController {

    private final static String PDF_TEMP_PATH = "template/pdf_template/pdf模板.pdf";
    private final static String FONT_CN = "font/simsun.ttc";

    /**
     * 注意事项:
     * 编辑PDF模板时，请参考 https://blog.csdn.net/weixin_43069862/article/details/102570657
     * @param response
     */
    @RequestMapping("/exportPDF")
    public void exportPDF(HttpServletResponse response) {
        try {
            List<PdfReader> readers = Lists.newArrayList();
            String[] str = {"1","2","3","4","5","6","7","8","9","10","1","2","3","4","5"};
            PdfReader pdfReader = ExportPDF.fillTemplate(str, PDF_TEMP_PATH, FONT_CN);
            readers.add(pdfReader);
            //合并导出的PDF
            ExportPDF.combinPdf(response, readers);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
