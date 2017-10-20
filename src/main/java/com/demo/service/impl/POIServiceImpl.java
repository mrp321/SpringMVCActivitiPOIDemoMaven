package com.demo.service.impl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.net.URLEncoder;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import com.demo.entity.User;
import com.demo.entity.UserInfo;
import com.demo.service.POIService;

/**
 * poi服务层
 * 
 * @author qchen
 * @date 2017-10-20
 * 
 */
@Service("poiService")
public class POIServiceImpl implements POIService {
	/**
	 * 转换为文件
	 * 
	 * @param userInfo
	 *            用户信息
	 * @throws FileNotFoundException
	 */
	@Override
	public void transfer2File(UserInfo userInfo) {
		String format = userInfo.getFormat();
		List<User> userList = userInfo.getUserList();
		File file = new File("C:\\员工信息一览." + format);
		OutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream(file);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		switch (format) {
		case "xls":
			this.transfer2Xls(userList, outputStream);
			break;
		case "xlsx":
			this.transfer2Xlsx(userList, outputStream);
			break;
		case "doc":
			this.transfer2Doc(userList, outputStream);
			break;
		case "docx":
			this.transfer2Docx(userList, outputStream);
			break;
		default:
			break;
		}
	}

	/**
	 * 转换为docx文件
	 * 
	 * @param userList
	 *            用户列表
	 * @param outputStream
	 *            输出流
	 * @return
	 */
	private void transfer2Docx(List<User> userList, OutputStream outputStream) {
		// XWPFParagraph：代表一个段落。
		// XWPFRun：代表具有相同属性的一段文本。
		// XWPFTable：代表一个表格。
		// XWPFTableRow：表格的一行。
		// XWPFTableCell：表格对应的一个单元格。
		// 获取list长度
		int userListLen = userList.size();
		XWPFDocument doc = new XWPFDocument();
		XWPFParagraph p1 = doc.createParagraph();

		// 生成 userListLen + 1 行，2列的表格
		XWPFTable table = doc.createTable(userListLen + 1, 2);
		CTTblPr tblPr = table.getCTTbl().getTblPr();
		tblPr.getTblW().setType(STTblWidth.DXA);
		tblPr.getTblW().setW(new BigInteger("7000"));

		// 设置上下左右四个方向的距离，可以将表格撑大
		table.setCellMargins(20, 20, 20, 20);
		// 表格
		List<XWPFTableCell> tableCells = table.getRow(0).getTableCells();
		tableCells.get(0).setText("用户id");
		tableCells.get(1).setText("密码");
		for (int i = 0; i < userListLen; i++) {
			User user = userList.get(i);
			tableCells = table.getRow(i + 1).getTableCells();
			tableCells.get(0).setText(user.getUserId());
			tableCells.get(1).setText(user.getPassword());
		}
		// 设置字体对齐方式
		p1.setAlignment(ParagraphAlignment.CENTER);
		p1.setVerticalAlignment(TextAlignment.TOP);
		// 第一页要使用p1所定义的属性
		XWPFRun r1 = p1.createRun();
		// 设置字体是否加粗
		r1.setBold(true);
		r1.setFontSize(20);
		// 设置使用何种字体
		r1.setFontFamily("Courier");
		// 设置上下两行之间的间距
		r1.setTextPosition(20);
		r1.setText("用户信息表");
		try {
			doc.write(outputStream);
			outputStream.close();
			doc.close();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {

		}
	}

	/**
	 * 转换为doc文件
	 * 
	 * @param userList
	 *            用户列表
	 * @param outputStream
	 *            输出流
	 * @return
	 */
	private void transfer2Doc(List<User> userList, OutputStream outputStream) {
		this.transfer2Docx(userList, outputStream);
	}

	/**
	 * 转换为xlsx文件
	 * 
	 * @param userList
	 *            用户列表
	 * @param outputStream
	 *            输出流
	 * @return
	 */
	private void transfer2Xlsx(List<User> userList, OutputStream outputStream) {
		SXSSFWorkbook workbook = new SXSSFWorkbook();
		this.genFile(workbook, userList, outputStream);

	}

	/**
	 * 转换为xls文件
	 * 
	 * @param userList
	 *            用户列表
	 * @param outputStream
	 *            输出流
	 * @return
	 */
	private void transfer2Xls(List<User> userList, OutputStream outputStream) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		this.genFile(workbook, userList, outputStream);
	}

	/**
	 * 生成本地文件
	 * 
	 * @param workbook
	 *            工作簿
	 * @param userList
	 *            用户列表
	 * @param outputStream
	 *            输出流
	 * @return
	 */
	private void genFile(Workbook workbook, List<User> userList,
			OutputStream outputStream) {
		Sheet sheet = workbook.createSheet("员工信息一览");
		// 样式
		CellStyle style = workbook.createCellStyle();
		// 字体
		Font font = workbook.createFont();
		Row row = null;
		Cell cell = null;
		// 设置边框
		style.setBorderBottom(BorderStyle.THIN); // 下边框
		style.setBorderLeft(BorderStyle.THIN);// 左边框
		style.setBorderTop(BorderStyle.THIN);// 上边框
		style.setBorderRight(BorderStyle.THIN);// 右边框
		font.setColor(HSSFFont.COLOR_NORMAL);
		font.setBold(false);
		style.setFont(font);
		// 为sheet生成第一行，用于放表头信息
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
		row = sheet.createRow(0);
		// 为第一行的第一个单元格的设置值
		cell = row.createCell(0);
		cell.setCellValue("用户信息一览");
		cell.setCellStyle(style);
		// 为sheet生成第二行，用于存放用户id和密码两个字段
		row = sheet.createRow(1);
		// 第二行第一个单元格存放用户id，
		cell = row.createCell(0);
		cell.setCellValue("用户id");
		cell.setCellStyle(style);
		// 第二行第二个单元格存放密码
		cell = row.createCell(1);
		cell.setCellValue("密码");
		cell.setCellStyle(style);
		// 为sheet生成剩余行，用于存放用户具体信息
		int len = userList.size();
		for (int i = 0; i < len; i++) {

			User user = userList.get(i);
			// 员工每添加一个，表格就再生成一行
			row = sheet.createRow(i + 2);
			// 这是用户id数据
			cell = row.createCell(0);
			cell.setCellValue(user.getUserId());
			cell.setCellStyle(style);
			// 这是用户密码
			cell = row.createCell(1);
			cell.setCellValue(user.getPassword());
			cell.setCellStyle(style);
		}
		try {
			workbook.write(outputStream);
			workbook.close();
			outputStream.flush();
			outputStream.close();

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 下载文件
	 * 
	 * @param format
	 *            格式
	 * @return
	 */

	@Override
	public ResponseEntity<byte[]> download(String format) {
		ResponseEntity<byte[]> entity = null;
		String origFileName = "员工信息一览.";
		try {
			String fileName = URLEncoder.encode(origFileName, "UTF-8");
			HttpHeaders headers = new HttpHeaders();
			headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
			headers.setContentDispositionFormData("attachment", fileName
					+ format);

			entity = new ResponseEntity<byte[]>(
					FileUtils.readFileToByteArray(new File("C:\\员工信息一览."
							+ format)), headers, HttpStatus.CREATED);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return entity;
	}

}
