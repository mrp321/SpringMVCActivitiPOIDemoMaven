package com.demo.controller;

import java.util.HashMap;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;

import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

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
	@RequestMapping(value = "transfer2File"/*, consumes = "application/json"*/)
	
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
}
