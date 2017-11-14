<?php if (!defined('THINK_PATH')) exit();?><!DOCTYPE html>
<html>
<head lang="en">
    <meta charset="UTF-8">
    <title></title>
</head>
<body>
<P><a href="<?php echo U('Exel/expUser');?>" >导出数据并生成excel</a></P><br/>
<form action="<?php echo U('Exel/impUser');?>" method="post" enctype="multipart/form-data">
    <input type="file" name="import"/>
    <input type="hidden" name="table" value="tablename"/>
    <input type="submit" value="导入"/>
</form>
</body>
</html>