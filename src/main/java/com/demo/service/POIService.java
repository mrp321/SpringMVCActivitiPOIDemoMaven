package com.demo.service;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.springframework.http.ResponseEntity;

import com.demo.entity.UserInfo;

public interface POIService {

	/**
	 * 转换为文件
	 * @param response 
	 * 
	 * @param request
	 * 
	 * @param company
	 *            公司
	 * @param format
	 *            格式
	 * @throws FileNotFoundException
	 */
	void transfer2File(UserInfo userInfo)
			throws FileNotFoundException;

	/**
	 * 下载文件
	 * 
	 * @return
	 * @throws IOException
	 */
	ResponseEntity<byte[]> download(String format);

}
