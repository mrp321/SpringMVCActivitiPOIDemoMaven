var uploadUrl = "/poi/uploadFile.do";
var transferUrl = "/poi/transferFile2Entity.do";

var page = new Vue({
	el : "#page",
	data : {
		userList : [],
		file : ""
	},
	methods : {
		upload : function() {
			var formData = new FormData($("#userForm")[0]);
			$.ajax({
				type : "post",
				contentType : false,
				url : uploadUrl,
				processData : false,
				data : formData,
				success : function(data) {
					console.log("upload:" + data.msg);
					if (data.success) {
						page.file = data.data[0];
					}
				}
			})
		},
		transfer : function() {
			$.ajax({
				type : "post",
				url : transferUrl,
				data : {
					file : page.file
				},
				success : function(data) {
					console.log("transfer:" + data.msg);
					if (data.success) {
						page.userList = data.data;
					}
				}
			})
		}
	}

})