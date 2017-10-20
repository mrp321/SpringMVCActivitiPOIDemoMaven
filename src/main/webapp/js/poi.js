var transfer2FileUrl = "/poi/transfer2File.do";
var downloadUrl = "/poi/download.do";
var page = new Vue({
	el : "#page",
	data : {
		userList : [],
		downloadUrlWithFormat : "",
	},
	mounted : function() {
		this.$nextTick(function() {
			page.initUser();
		})
	},
	methods : {
		initUser : function() {
			page.userList = [];
			console.log("initUser:userList=");
			console.log(page.userList);
		},
		delUser : function(user) {
			// 移除指定元素
			page.userList.splice($.inArray(user, page.userList), 1);
			console.log("delUser:userList=");
			console.log(page.userList);
		},
		addUser : function() {
			// 获取添加输入框中内容
			var userIdAdd = $("#userIdAdd").val();
			var passwordAdd = $("#passwordAdd").val();
			// 定义添加的用户实体
			var user = new Object();
			user.userId = userIdAdd;
			user.password = passwordAdd;
			// 在userList中添加该用户实体
			page.userList.push(user);
			console.log(page.userList);
			// 添加用户完后，清除添加用户输入框中的内容
			$("#userIdAdd").val("");
			$("#passwordAdd").val("");
			console.log("addUser:userList=");
			console.log(page.userList);
		},
		// 转换
		transfer : function() {
			console.log("transfer:userList=");
			console.log(page.userList);
			var userInfo = new Object();
			userInfo.userList = page.userList;
			userInfo.format = $("#format").val();
			// 对象转为json字符串
			userInfo = JSON.stringify(userInfo);
			$.ajax({
				type : "post",
				url : transfer2FileUrl,
				async : false, 
				data : userInfo,
				contentType : "application/json",
				success : function(data) {
					console.log("transfer result=" + data.msg);
					if (data.success) {
						page.download();
					}
				}
			})
		},
		// 下载
		download : function() {
			var format = $("#format").val();
			page.downloadUrlWithFormat = downloadUrl + "?format=" + format;
		}
	}
})
