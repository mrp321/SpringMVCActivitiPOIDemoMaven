package com.demo.service;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import javax.servlet.http.HttpServletRequest;

import org.springframework.http.ResponseEntity;

import com.demo.entity.User;
import com.demo.entity.UserInfo;

public interface POIService {

	/**
	 * 转换为文件
	 * 
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
	void transfer2File(UserInfo userInfo) throws FileNotFoundException;

	/**
	 * 下载文件
	 * 
	 * @return
	 * @throws IOException
	 */
	ResponseEntity<byte[]> download(String format);

	/**
	 * 将文件转换为实体
	 * 
	 * @param request
	 * 
	 * @param request
	 * @return
	 */
	List<User> transferFile2Entity(HttpServletRequest request, String file);

	/**
	 * 文件上传
	 * 
	 * @param request
	 * @return
	 */
	List<String> uploadFile(HttpServletRequest request);

}
