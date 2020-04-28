package com.example.demoio.core.util;

import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.List;
import java.util.NoSuchElementException;

/**
 * 导出execl
 *
 * @author lixingxing
 * @date 2019年9月23日 10点29分
 */
@Slf4j
public class ExcelUtil {

	/**
	 * 导出excel
	 *
	 * @param title    导出表的标题
	 * @param rowsName 导出表的列名
	 * @param dataList 需要导出的数据
	 * @param fileName 生成excel文件的文件名
	 * @param response
	 */
	public static void exportExcel(String title, String[] rowsName, List<Object[]> dataList, String fileName, HttpServletResponse response, HSSFSheet sheet, HSSFWorkbook workbook) throws Exception {
		OutputStream output = response.getOutputStream();
		response.reset();
		response.setContentType("application/octet-stream; charset=UTF-8");
		response.setHeader("Content-disposition",
				"attachment; filename=" + URLEncoder.encode(fileName + ".xls"));
		try {
			HSSFRow rowm = sheet.createRow(0);  // 产生表格标题行
			HSSFCell cellTiltle = rowm.createCell(0);   //创建表格标题列
			// sheet样式定义;    getColumnTopStyle();    getStyle()均为自定义方法 --在下面,可扩展
			HSSFCellStyle columnTopStyle = getColumnTopStyle(workbook,15);// 获取列头样式对象
			HSSFCellStyle style = getStyle(workbook); // 获取单元格样式对象
			//合并表格标题行，合并列数为列名的长度,第一个0为起始行号，第二个1为终止行号，第三个0为起始列好，第四个参数为终止列号
			sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, (rowsName.length - 1)));
			cellTiltle.setCellStyle(columnTopStyle);    //设置标题行样式
			cellTiltle.setCellValue(title);     //设置标题行值
			int columnNum = rowsName.length;     // 定义所需列数
			HSSFRow rowRowName = sheet.createRow(2); // 在索引2的位置创建行(最顶端的行开始的第二行)
			// 将列头设置到sheet的单元格中
			for (int n = 0; n < columnNum; n++) {
				HSSFCell cellRowName = rowRowName.createCell(n); // 创建列头对应个数的单元格
				cellRowName.setCellType(CellType.STRING); // 设置列头单元格的数据类型
				HSSFRichTextString text = new HSSFRichTextString(rowsName[n]);
				cellRowName.setCellValue(text); // 设置列头单元格的值
				cellRowName.setCellStyle(columnTopStyle); // 设置列头单元格样式
				sheet.setColumnWidth(n,5000); //设置宽度
			}

			// 将查询出的数据设置到sheet对应的单元格中
			for (int i = 0; i < dataList.size(); i++) {
				Object[] obj = dataList.get(i);   // 遍历每个对象
				HSSFRow row = sheet.createRow(i + 3);   // 创建所需的行数
				for (int j = 0; j < obj.length; j++) {
					HSSFCell cell = row.createCell(j, CellType.STRING);
					if (!"".equals(obj[j]) && obj[j] != null) {
						cell.setCellValue(obj[j].toString()); // 设置单元格的值
					} else {
						cell.setCellValue(""); // 设置单元格的值
					}
					cell.setCellStyle(style); // 设置单元格样式
				}
			}
			workbook.write(output);
			output.flush();
			response.flushBuffer();
		} catch (Exception e) {
			e.printStackTrace();
		}
		close(output);
	}


	/*
	 * 列头单元格样式
	 * @param i 字体大小
	 */
	private static HSSFCellStyle getColumnTopStyle(HSSFWorkbook workbook , int i) {

		// 设置字体
		HSSFFont font = workbook.createFont();
		// 设置字体大小
		font.setFontHeightInPoints((short) i);
		// 字体加粗
		font.setBold(true);
		// 设置字体名字
		font.setFontName("Courier New");
		// 设置样式;
		HSSFCellStyle style = workbook.createCellStyle();
		// 设置底边框;
		style.setBorderBottom(BorderStyle.THIN);
		// 设置底边框颜色;
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		// 设置左边框;
		style.setBorderLeft(BorderStyle.THIN);
		// 设置左边框颜色;
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		// 设置右边框;
		style.setBorderRight(BorderStyle.THIN);
		// 设置右边框颜色;
		style.setRightBorderColor(IndexedColors.BLACK.index);
		// 设置顶边框;
		style.setBorderTop(BorderStyle.THIN);
		// 设置顶边框颜色;
		style.setTopBorderColor(IndexedColors.BLACK.index);
		// 在样式用应用设置的字体;
		style.setFont(font);
		// 设置自动换行;
		style.setWrapText(false);
		// 设置水平对齐的样式为居中对齐;
		style.setAlignment(HorizontalAlignment.CENTER);
		// 设置垂直对齐的样式为居中对齐;
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		return style;

	}

	/*
	 * 列数据信息单元格样式
	 */
	private static HSSFCellStyle getStyle(HSSFWorkbook workbook) {
		// 设置字体
		HSSFFont font = workbook.createFont();
		// 设置字体大小
		font.setFontHeightInPoints((short) 10);
		// 字体加粗
		//	font.setBold(true);
		// 设置字体名字
		font.setFontName("Courier New");
		// 设置样式;
		HSSFCellStyle style = workbook.createCellStyle();
		// 设置底边框;
		style.setBorderBottom(BorderStyle.THIN);
		// 设置底边框颜色;
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		// 设置左边框;
		style.setBorderLeft(BorderStyle.THIN);
		// 设置左边框颜色;
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		// 设置右边框;
		style.setBorderRight(BorderStyle.THIN);
		// 设置右边框颜色;
		style.setRightBorderColor(IndexedColors.BLACK.index);
		// 设置顶边框;
		style.setBorderTop(BorderStyle.THIN);
		// 设置顶边框颜色;
		style.setTopBorderColor(IndexedColors.BLACK.index);
		// 在样式用应用设置的字体;
		style.setFont(font);
		// 设置自动换行;
		style.setWrapText(false);
		// 设置水平对齐的样式为居中对齐;
		style.setAlignment(HorizontalAlignment.CENTER);
		// 设置垂直对齐的样式为居中对齐;
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		return style;
	}

	/**
	 * 关闭输出流
	 *
	 * @param os
	 */
	private static void close(OutputStream os) {
		if (os != null) {
			try {
				os.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}


	/**
	 * 导入
	 * @param file 文件
	 * @param titleRows 标题是前几行
	 * @param headerRows 头部是前几行
	 * @param pojoClass 接收的对象
	 * @param <T>
	 * @return
	 */
	public static <T> List<T> importExcel(MultipartFile file, Integer titleRows, Integer headerRows, Class<T> pojoClass){
		if (file == null){
			return null;
		}
		ImportParams params = new ImportParams();
		params.setTitleRows(titleRows);
		params.setHeadRows(headerRows);
		List<T> list = null;
		try {
			list = ExcelImportUtil.importExcel(file.getInputStream(), pojoClass, params);
		}catch (NoSuchElementException e){
			throw e;
		} catch (Exception e) {
			e.printStackTrace();
			log.error("[monitor][表单功能]", e);
		}
		return list;
	}
}

