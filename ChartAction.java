package com.success.action;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.rmi.RemoteException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import jxl.Cell;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.struts2.ServletActionContext;

import com.success.service.ChartService;
import com.success.util.BaseAction;
import com.success.util.EduceToExcel;
import com.success.util.ExportExcel;

public class ChartAction extends BaseAction {
	HttpServletRequest request = ServletActionContext.getRequest();

	private String bgCode;

	public String getBgCode() {
		return bgCode;
	}

	public void setBgCode(String bgCode) {
		bgCode = request.getParameter("bgCode");
		this.bgCode = bgCode;
	}

	String data = "";
	ChartService cs = null;

	public void setTheme() {
		cs = new ChartService();
		try {
			cs.saveTheme((String) ServletActionContext.getRequest()
					.getSession().getAttribute("username"), bgCode);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void getWrylb() {
		cs = new ChartService();
		try {
			data = cs.getWrylb(4);
			// //System.out.println(data);
			send(data);
		} catch (Throwable e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void getFspw() {
		cs = new ChartService();
		try {
			data = cs.getFspwYears(3);
			 System.out.println("废水"+data);
			send(data);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void getFspwExcel() {
		ExportExcel exportExcel = new ExportExcel();
		String fileName = "工业废水排放统计";// 文件路径
		JSONArray jsonData = getJsonData("fspw");
		String[] columnName = { "年份", "废水(亿吨)", "氨氮(万吨)", "COD(万吨)" };
		String[] columnType = { "年度", "废水", "氨氮", "COD" };
		exportExcel
				.Export(fileName,jsonData, columnName, columnType);
	}

	public void getFqpw() {
		cs = new ChartService();
		try {
			data = cs.getFqpwYears(3);
			 System.out.println("废气"+data);
			send(data);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void getFqpwExcel() {
		ExportExcel exportExcel = new ExportExcel();
		String fileName = "工业废气排放统计";// 文件路径
		JSONArray jsonData = getJsonData("fqpw");
		String[] columnName = { "年份", "废气(亿吨)", "二氧化硫(万吨)", "氮氧化物(万吨)",
				"烟尘(万吨)" };
		String[] columnType = { "年度", "废气", "SO2", "NOx", "烟尘" };
		exportExcel
				.Export(fileName,jsonData, columnName, columnType);
	}

	public void getJsxm() {
		cs = new ChartService();
		try {
			data = cs.getJsxm(3);
			// //System.out.println(data);
			send(data);
		} catch (RemoteException e) {
			e.printStackTrace();
		}
	}

	public void getXzcf() {
		cs = new ChartService();
		try {
			data = cs.getXzcf(3);
			// //System.out.println(data);
			send(data);
		} catch (RemoteException e) {
			e.printStackTrace();
		}
	}

	public void getXzcfExcel() {
		ExportExcel exportExcel = new ExportExcel();
		String fileName = "行政处罚统计";// 文件路径
		JSONArray jsonData = getJsonData("xzcf");
		String[] columnName = { "年份", "案件个数", "总罚款金额(万元)" };
		String[] columnType = { "nd", "gs", "fkze" };
		exportExcel
				.Export(fileName,jsonData, columnName, columnType);
	}

	// 环统企业个数统计
	public void getHt() {
		cs = new ChartService();
		try {
			data = cs.getHt(2);
			// System.out.println(data);
			send(data);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void getHtExcel() throws UnsupportedEncodingException {
//		HttpServletResponse response = ServletActionContext.getResponse();
//		response.setContentType("application/vnd.ms-excel; charset=utf-8");
//		String fileName1 = "环统企业个数统计.xlsx";
//		String fileName2 = new String(fileName1.getBytes(), "iso_8859_1");// 设置文件名称的编码格式
//		response.setHeader("Content-Disposition", "attachment;filename="
//				+ fileName2);
//		response.setHeader("Pragma", "no-cache");//no-cache指示请求或响应消息不能缓存
//		response.setHeader("Cache-Control", "no-cache");//no-cache指示请求或响应消息不能缓存 
//		response.setDateHeader("Expires", 0);
//		response.setCharacterEncoding("utf-8");
		ExportExcel exportExcel = new ExportExcel();
		String fileName = "环统企业个数统计";// 文件路径
		JSONArray jsonData = getJsonData("ht");
		String[] columnName = { "年份", "环统企业个数" };
		String[] columnType = { "nd", "qys" };
		exportExcel.Export(fileName,  jsonData, columnName, columnType);
	}

	public String getHjzl() {
		cs = new ChartService();
		try {
			data = cs.getHjzl();
			ServletActionContext.getRequest().getSession()
					.setAttribute("hzjlData", data);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return "hjzl";
	}

	/**
	 * 读取json数据并解析
	 * 
	 * @return
	 * @throws IOException
	 */
	public JSONArray getJsonData(String name) {
		cs = new ChartService();
		JSONArray jsonArray = null;
		try {
			if (name.equals("ht")) {
				String json = cs.getHt(2);
				JSONObject jsonObject = JSONObject.fromObject(json);// 把String转成JSONObject类型
				if (jsonObject != null) {
					jsonArray = jsonObject.getJSONArray("ht");
				}
			} else if (name.equals("xzcf")) {
				String json = cs.getXzcf(3);
				JSONObject jsonObject = JSONObject.fromObject(json);// 把String转成JSONObject类型
				if (jsonObject != null) {
					jsonArray = jsonObject.getJSONArray("xzcf");
				}
			} else if (name.equals("fspw")) {
				String json = cs.getFspwYears(3);
				JSONObject jsonObject = JSONObject.fromObject(json);// 把String转成JSONObject类型
				if (jsonObject != null) {
					jsonArray = jsonObject.getJSONArray("fspw");
				}

			} else if (name.equals("fqpw")) {
				String json = cs.getFqpwYears(3);
				JSONObject jsonObject = JSONObject.fromObject(json);// 把String转成JSONObject类型
				if (jsonObject != null) {
					jsonArray = jsonObject.getJSONArray("fqpw");
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			jsonArray = null;
		}
		return jsonArray;
	}

	public static void main(String[] args) {
		String json = "{\"ht\":[{\"nd\":2016,\"qys\":29},{\"nd\":2017,\"qys\":18},{\"nd\":2018,\"qys\":15},{\"nd\":2019,\"qys\":3}]}";
		JSONObject jsonObject = JSONObject.fromObject(json);// 把String转成JSONObject类型
		JSONArray jsonData = jsonObject.getJSONArray("ht");
		ExportExcel exportExcel = new ExportExcel();
		String fileName = "环统企业统计";// 文件路径
		String[] columnName = { "年份", "环统企业个数" };
		String[] columnType = { "nd", "qys" };
		exportExcel.Export(fileName, jsonData, columnName, columnType);

		// String json =
		// "{\"fqpw\":[{\"年度\":\"2015\",\"废气\":0.63754181012121,\"SO2\":15.05296234,\"NOx\":4.6026899704,\"烟尘\":152.8484173856},{\"年度\":\"2016\",\"废气\":0.40590049706780,\"SO2\":10.7758623156,\"NOx\":4.5857972420,\"烟尘\":166.2026142188},{\"年度\":\"2017\",\"废气\":0.64539096972194,\"SO2\":14.8299112852,\"NOx\":2.39999167,\"烟尘\":183.7349297124}]}";
		// JSONObject jsonObject = JSONObject.fromObject(json);//
		// 把String转成JSONObject类型
		// JSONArray jsonData = jsonObject.getJSONArray("fqpw");
		// ExportExcel exportExcel = new ExportExcel();
		// String fileName = "C:/Users/Administrator/Desktop/工业废气排放统计.xlsx";//
		// 文件路径
		// String pageName = "工业废气排放统计";
		// // JSONArray jsonData = getJsonData("fqpw");
		// String[] columnName = { "年份", "废气(亿吨)", "二氧化硫(万吨)",
		// "氮氧化物(万吨)","烟尘(万吨)" };
		// String[] columnType = { "年度", "废气", "SO2", "NOx", "烟尘" };
		// exportExcel.Export(fileName, pageName, jsonData, columnName,
		// columnType);
	}
}
