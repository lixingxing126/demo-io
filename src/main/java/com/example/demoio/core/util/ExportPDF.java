package com.example.demoio.core.util;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import lombok.extern.log4j.Log4j;
import org.apache.commons.collections4.CollectionUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.Iterator;
import java.util.List;

@Log4j
public class ExportPDF {

	/**
	 * 使用模板生成PDF
	 *
	 * @param str          充填表格的数据
	 * @param templatePath 模板路径
	 * @param fontCn       中文包
	 */
	public static PdfReader fillTemplate(String[] str, String templatePath, String fontCn) {
		PdfReader reader;
		ByteArrayOutputStream bos;
		PdfStamper stamper;
		try {
			reader = new PdfReader(templatePath);// 读取pdf模板
			bos = new ByteArrayOutputStream();
			stamper = new PdfStamper(reader, bos);
			AcroFields form = stamper.getAcroFields();
			int i = 0;
			Iterator<String> it = form.getFields().keySet().iterator();
			//设置中文字体，如果不设置部分中文不显示。
			ExportPDF exportPDF = new ExportPDF();
			//BaseFont bf  = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
			BaseFont bf = BaseFont.createFont(fontCn + ",0" ,BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
//			form.setSubstitutionFonts(Lists.newArrayList(bf));
			form.addSubstitutionFont(bf);
			while (it.hasNext()) {
				String name = it.next().toString();
				form.setField(name, str[i++]);
			}
			stamper.setFormFlattening(true);// 如果为false那么生成的PDF文件还能编辑，一定要设为true
			stamper.close();
			Document doc = new Document();
			doc.open();
			PdfReader pdfReader = new PdfReader(bos.toByteArray());
			doc.close();
			return pdfReader;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}


	/**
	 * 合并 pdf
	 *
	 * @param readers  所有要合并的PDF的集合
	 * @param response
	 * @throws Exception
	 */
	public static void combinPdf(HttpServletResponse response, List<PdfReader> readers) {
		String targetFilename = "今日案件.pdf";
		response.setContentType("application/octet-stream; charset=UTF-8");
		response.setHeader("Content-disposition",
				"attachment; filename=" + URLEncoder.encode(targetFilename));

		try {
			if (CollectionUtils.isNotEmpty(readers)) {
				Document document = new Document();
				PdfWriter writer = PdfWriter.getInstance(document, response.getOutputStream());
				document.open();
				PdfContentByte cb = writer.getDirectContent();
				int pageOfCurrentReaderPDF = 0;
				Iterator<PdfReader> iteratorPDFReader = readers.iterator();
				while (iteratorPDFReader.hasNext()) {
					PdfReader pdfReader = iteratorPDFReader.next();
					while (pageOfCurrentReaderPDF < pdfReader.getNumberOfPages()) {
						document.newPage();
						pageOfCurrentReaderPDF++;
						PdfImportedPage page = writer.getImportedPage(pdfReader, pageOfCurrentReaderPDF);
						cb.addTemplate(page, 0, 0);
					}
					pageOfCurrentReaderPDF = 0;
				}
				document.close();
				writer.close();
			} else {
				Document doc = new Document(PageSize.B5, 0, 0, 0, 0);
				PdfWriter writer = PdfWriter.getInstance(doc, response.getOutputStream());
				doc.open();
				BaseFont bfChinese = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
				Font FontChinese1 = new Font(bfChinese, 18, Font.BOLD);
				Paragraph t = new Paragraph("暂无数据", FontChinese1);
				t.setAlignment(Paragraph.ALIGN_CENTER);
				doc.add(t);
				doc.close();
				writer.close();
			}
		} catch (IOException | DocumentException e) {
			e.printStackTrace();
		}
	}

}
