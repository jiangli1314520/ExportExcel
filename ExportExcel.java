package com.success.util;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.vfs.FileUtil;
import org.apache.struts2.ServletActionContext;

import jxl.CellView;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import com.success.service.ChartService;

public class ExportExcel {
	/**
	 * @param fileName
	 *            路径
	 * @param pageName
	 *            页名
	 * @param jsonData
	 *            json数据
	 * @param columnName
	 *            列名
	 * @param columnType
	 *            json数据key 只存储一级json,第一页
	 */
	public void Export(String fileName, JSONArray jsonData,
			String[] columnName, String[] columnType) {

		HttpServletResponse response = ServletActionContext.getResponse();
		try {
			WritableWorkbook wwb = null;
			File file = new File("C:/Users/Administrator/Downloads/" + fileName+"副本.xlsx");
			// 以fileName为文件名来创建一个Workbook
			wwb = Workbook.createWorkbook(file);
			// 创建工作表
			WritableSheet ws = wwb.createSheet(fileName, 0);
			// 设置字体
			WritableFont font0 = new WritableFont(WritableFont.ARIAL, 14,
					WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE,
					Colour.BLACK);
			WritableCellFormat cellFormat0 = new WritableCellFormat(font0);
			// 设置文字居中对齐方式
			cellFormat0.setAlignment(Alignment.CENTRE);
			// 设置垂直居中
			cellFormat0.setVerticalAlignment(VerticalAlignment.CENTRE);
			// 合并单元格
			/**
			 * @param arg0
			 *            第一个参数：要合并的单元格最左上角的列号
			 * @param arg1
			 *            第二个参数：要合并的单元格最左上角的行号
			 * @param arg2
			 *            第三个参数：要合并的单元格最右角的列号
			 * @param arg3
			 *            第四个参数：要合并的单元格最右下角的行号
			 * @param columnType
			 */
			ws.mergeCells(0, 0, columnName.length - 1, 0);
			Label label0 = new Label(0, 0, fileName, cellFormat0);// 表示第
			ws.addCell(label0);
			// 设置字体
			/**
			 * @param WritableFont
			 *            .createFont("宋体")：设置字体为宋体
			 * @param 10：设置字体大小
			 * @param WritableFont
			 *            .BOLD:设置字体加粗（BOLD：加粗 NO_BOLD：不加粗）
			 * @param false：设置非斜体
			 * @param UnderlineStyle
			 *            .NO_UNDERLINE：没有下划线
			 * @param Colour
			 *            .BLACK：字体颜色
			 */
			WritableFont font1 = new WritableFont(WritableFont.ARIAL, 10,
					WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE,
					Colour.BLACK);
			WritableCellFormat cellFormat1 = new WritableCellFormat(font1);
			// 设置文字居中对齐方式
			cellFormat1.setAlignment(Alignment.CENTRE);
			// 设置垂直居中
			cellFormat1.setVerticalAlignment(VerticalAlignment.CENTRE);

			// 给sheet电子版中所有的列设置默认的列的宽度
			ws.getSettings().setDefaultColumnWidth(17);
			// 根据内容自动设置列宽
			// CellView cellView = new CellView();
			// cellView.setAutosize(true); // 设置自动大小

			// 要插入到的Excel表格的行号，默认从0开始
			for (int i = 0; i < columnName.length; i++) {
				// ws.setColumnView(i, cellView); // 根据内容自动设置列宽
				// 样式end
				Label label = new Label(i, 1, columnName[i], cellFormat1);// 表示第
				ws.addCell(label);
			}
			for (int i = 0; i < jsonData.size(); i++) {
				JSONObject firstModularObject = jsonData.getJSONObject(i);
				for (int j = 0; j < columnType.length; j++) {
					Label label = new Label(j, i + 2,
							firstModularObject.getString(columnType[j]),
							cellFormat1);
					ws.addCell(label);
				}
			}
			// 写进文档
			wwb.write();
			// 关闭Excel工作簿对象
			wwb.close();
			// Excel下载
			// 设置excel文件名
			response.setHeader("Content-Disposition", "attachment; filename="+ URLEncoder.encode(fileName, "utf-8")+URLEncoder.encode(".xlsx", "utf-8"));
			response.setHeader("Pragma", "no-cache");// no-cache指示请求或响应消息不能缓存
			response.setHeader("Cache-Control", "no-cache");// no-cache指示请求或响应消息不能缓存
			response.setDateHeader("Expires", 0);
			response.setCharacterEncoding("utf-8");
			// 4.获取要下载的文件输入流
			InputStream in = new FileInputStream(file);
			int len = 0;
			// 5.创建数据缓冲区
			byte[] buffer = new byte[1024];
			// 6.通过response对象获取OutputStream流
			OutputStream out = response.getOutputStream();
			// 7.将FileInputStream流写入到buffer缓冲区
			while ((len = in.read(buffer)) > 0) {
				// 8.使用OutputStream将缓冲区的数据输出到客户端浏览器
				out.write(buffer, 0, len);
			}
			in.close();
			out.flush();
			response.flushBuffer();// 不可少
			out.close();
			file.delete();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
