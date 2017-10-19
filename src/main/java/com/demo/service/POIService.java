package com.demo.service;

import java.io.FileNotFoundException;

import com.demo.entity.UserInfo;

public interface POIService {

	/**
	 * 转换为文件
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

}
