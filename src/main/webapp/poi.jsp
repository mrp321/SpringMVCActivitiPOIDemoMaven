<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>poi演示</title>
</head>
<body>
	<div id="page">
		<table>
			<thead>
				<tr>
					<th>用户id</th>
					<th>密码</th>
					<th>操作</th>
				</tr>
			</thead>
			<tbody>
				<tr v-for="(user,index) in userList">
					<td><input type="text" id="'userId' + index"
						:value="user.userId"></td>
					<td><input type="text" id="'password' + index"
						:value="user.password"></td>
					<td><input type="button" value="删除" @click="delUser(user)"></td>
				</tr>
				<tr>
					<td><input type="text" id="userIdAdd"></td>
					<td><input type="text" id="passwordAdd"></td>
					<td><input type="button" value="添加" @click="addUser()"></td>
				</tr>
			</tbody>
		</table>
		<div>
			<select id="format">
				<option value="xls" selected="selected">xls</option>
				<option value="xlsx">xlsx</option>
<!-- 				<option value="doc">doc</option>
				<option value="docx">docx</option> -->
			</select> <input type="button" value="生成" @click="transfer()"> <a :href="downloadUrlWithFormat">下载文件</a>
		</div>
	</div>
	<script type="text/javascript" src="js/jquery-1.10.1.min.js"></script>
	<script type="text/javascript" src="js/vue.js"></script>
	<script type="text/javascript" src="js/poi.js"></script>
</body>
</html>