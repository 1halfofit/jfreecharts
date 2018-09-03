<%--
  Created by IntelliJ IDEA.
  User: ibm
  Date: 2018/8/28
  Time: 16:38
  To change this template use File | Settings | File Templates.
--%>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title>Excel导入导出</title>
</head>
<body>
<p>上传Excel:</p>
<form action="/importexcel.do" method="post" enctype="multipart/form-data">
    <input type="file" name="file" /> <input type="submit" value="Submit" /></form>


<p>导出Excel:</p>
<a href="/exceldownload.do">导出EXCEL</a>

</body>
</html>
