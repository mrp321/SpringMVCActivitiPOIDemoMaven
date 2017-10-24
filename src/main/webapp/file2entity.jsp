<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>使用poi将文件内容转换为实体</title>
</head>
<body>
	<div id="page">
		<div>
			<form name="userForm" id="userForm" action=""
				enctype="multipart/form-data" method="post">

				请选择文件：<input type="file" name="file" id="file"><br> <input
					type="button" value="上传" @click="upload"><br> <input
					type="button" value="显示转换结果" @click="transfer()">
			</form>
		</div>
		<div v-if="userList != null && userList != [] && userList != ''">
			<table border="1" cellpadding="1" cellspacing="1">
				<thead>
					<tr>
						<th>用户id</th>
						<th>密码</th>
					</tr>
				</thead>
				<tbody>
					<tr v-for="(user,index) in userList">
						<td><span>{{ user.userId }}</span></td>
						<td><span>{{ user.password }}</span></td>
					</tr>
				</tbody>
			</table>
		</div>
	</div>
	<script type="text/javascript" src="js/jquery-1.10.1.min.js"></script>
	<script type="text/javascript" src="js/vue.js"></script>
	<script type="text/javascript" src="js/transfer2entity.js"></script>
</body>
</html>