package com.demo.service.impl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Service;

import com.demo.entity.User;
import com.demo.entity.UserInfo;
import com.demo.service.POIService;

@Service("poiService")
public class POIServiceImpl implements POIService {
	/**
	 * 转换为文件
	 * 
	 * @param company
	 *            公司
	 * @param format
	 *            格式
	 * @throws FileNotFoundException
	 */
	@Override
	public void transfer2File(UserInfo userInfo) {
		String format = userInfo.getFormat();
		List<User> userList = userInfo.getUserList();
		File file = new File("C:\\1." + format);
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

	private void transfer2Docx(List<User> userList, OutputStream outputStream) {
		// TODO Auto-generated method stub

	}

	private void transfer2Doc(List<User> userList, OutputStream outputStream) {
		// TODO Auto-generated method stub

	}

	private void transfer2Xlsx(List<User> userList, OutputStream outputStream) {
		// TODO Auto-generated method stub

	}

	private void transfer2Xls(List<User> userList, OutputStream outputStream) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("员工信息一览");
		HSSFRow row = null;
		HSSFCell cell = null;
		// 为sheet生成第一行，用于放表头信息
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
		row = sheet.createRow(0);
		// 为第一行的第一个单元格的设置值
		cell = row.createCell(0);
		cell.setCellValue("用户信息一览");
		// 为sheet生成第二行，用于存放用户id和密码两个字段
		row = sheet.createRow(1);
		// 第二行第一个单元格存放用户id，
		cell = row.createCell(0);
		cell.setCellValue("用户id");
		// 第二行第二个单元格存放密码
		cell = row.createCell(1);
		cell.setCellValue("密码");
		// 为sheet生成剩余行，用于存放用户具体信息
		int len = userList.size();
		for (int i = 0; i < len; i++) {
			User user = userList.get(i);
			// 员工每添加一个，表格就再生成一行
			row = sheet.createRow(i + 2);
			// 这是用户id数据
			cell = row.createCell(0);
			cell.setCellValue(user.getUserId());
			// 这是用户密码
			cell = row.createCell(1);
			cell.setCellValue(user.getPassword());
		}
		try {
			workbook.write(outputStream);
			outputStream.flush();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
