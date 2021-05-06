<?php

header("content-type:text/html;charset=utf-8");

exec("python3 /var/www/html/QualityCtrl/No1FileTool/preForB.py 2>&1",$out,$ret);
print_r($out);
print_r($ret);


header("Location: http://47.114.178.105/QualityCtrl/No1FileTool/upload_confirm.html");
?>
