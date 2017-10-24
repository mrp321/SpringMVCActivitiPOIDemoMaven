package com.demo.controller;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;

import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.demo.entity.User;
import com.demo.entity.UserInfo;
import com.demo.service.POIService;
import com.xiaoleilu.hutool.util.StrUtil;

/**
 * POI测试controller
 * 
 * @date 2017-10-19
 * @author qchen
 * 
 */
@Controller
@RequestMapping("/poi/")
public class POIController extends CommController {
	/**
	 * 日志
	 */
	private static final Logger lg = Logger.getLogger(POIController.class);

	@Autowired
	private POIService poiService;

	/**
	 * 转为文件
	 * 
	 * @param request
	 * @param company
	 *            公司
	 * @param format
	 *            文件格式
	 * @return
	 */
	@RequestMapping(value = "transfer2File"/* , consumes = "application/json" */)
	@ResponseBody
	public Map<String, Object> transfer2File(HttpServletRequest request,
			@RequestBody UserInfo userInfo) {
		Map<String, Object> map = new HashMap<String, Object>(16);
		if (userInfo != null) {
			lg.info("转换开始，传入参数-> userInfo=" + userInfo);
			try {
				this.poiService.transfer2File(userInfo);
				map = this.getSuccessMap("转换成功");
			} catch (Exception e) {
				map = this.getFailureMap("转换失败，信息：" + e.getMessage());
				e.printStackTrace();
			}
		} else {
			map = this.getEmptyParamsMap();
		}
		return map;
	}

	/**
	 * 下载文件
	 * 
	 * @return
	 * @throws Exception
	 * @throws IOException
	 */
	@RequestMapping("download")
	public ResponseEntity<byte[]> download(HttpServletRequest request,
			String format) throws Exception {
		if (StrUtil.isNotEmpty(format)) {
			ResponseEntity<byte[]> entity = null;
			entity = this.poiService.download(format);
			return entity;
		} else {
			throw new Exception("传入参数不能为空");
		}

	}

	/**
	 * 文件上传
	 * 
	 * @param request
	 * @return
	 */
	@RequestMapping("uploadFile")
	@ResponseBody
	public Map<String, Object> uploadFile(HttpServletRequest request) {
		Map<String, Object> map = new HashMap<>();
		try {
			List<String> fileList = this.poiService.uploadFile(request);
			map = this.getReturnMap(true, fileList, "上传成功");
		} catch (Exception e) {
			map = this.getFailureMap("上传失败，信息：" + e.getMessage());
		}
		return map;
	}

	/**
	 * 将文件转换为实体
	 * 
	 * @param request
	 * @param file
	 *            文件
	 * @return
	 */
	@RequestMapping("transferFile2Entity")
	@ResponseBody
	public Map<String, Object> transferFile2Entity(HttpServletRequest request,
			String file) {
		Map<String, Object> map = new HashMap<>();
		try {
			List<User> userList = this.poiService.transferFile2Entity(request,
					file);
			map = this.getReturnMap(true, userList, "转换成功");
		} catch (Exception e) {
			map = this.getFailureMap("转换失败，信息：" + e.getMessage());
		}
		return map;
	}
}
