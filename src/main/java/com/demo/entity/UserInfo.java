package com.demo.entity;

import java.util.List;

/**
 * 用户信息实体
 * 
 * @author qchen
 * @date 2017-10-19
 * 
 */
public class UserInfo {
	/** 用户列表 */
	private List<User> userList;

	public List<User> getUserList() {
		return userList;
	}

	public void setUserList(List<User> userList) {
		this.userList = userList;
	}

	@Override
	public String toString() {
		return "UserInfo [userList=" + userList + "]";
	}
}
