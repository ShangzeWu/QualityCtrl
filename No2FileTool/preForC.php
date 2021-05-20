<?php

header("content-type:text/html;charset=utf-8");

exec("python3 /var/www/html/QualityCtrl/No2FileTool/preForC.py 2>&1",$out,$ret);
print_r($out);
print_r($ret);
exec("python3 /var/www/html/QualityCtrl/No2FileTool/rmvoid.py 2>&1",$out1,$ret1);
print_r($out1);
print_r($ret1);
exec("python3 /var/www/html/QualityCtrl/No2FileTool/CtoTemp.py 2>&1",$out2,$ret2);
print_r($out2);
print_r($ret2);


#$shell="python3 preForA.py";
#$a=exec($shell."2>error.txt",$array,$ret);
#print_r($a);
#echo $return_var;


header("Location: http://47.114.178.105/QualityCtrl/No2FileTool/resultC");
?>
