<?php

$allowedExts = array("xlsx");
$temp = explode(".", $_FILES["file"]["name"]);  //以.作为分隔符
echo $_FILES["file"]["size"]."kb"."<br>";
#echo "请注意，文件大小不能超过2000kb"."<br>";
$extension = end($temp);     // 获取文件后缀名
if (($extension == "xlsx")
#&& ($_FILES["file"]["size"] < 2048000)   // 小于 2000 kb
&& in_array($extension, $allowedExts))
{
    if ($_FILES["file"]["error"] > 0)
    {
        echo "错误：: " . $_FILES["file"]["error"] . "<br>";
    }
    else
    {
        echo "上传文件名: " . $_FILES["file"]["name"] . "<br>";
        echo "文件类型: " . $_FILES["file"]["type"] . "<br>";
        echo "文件大小: " . ($_FILES["file"]["size"] / 1024) . " kB<br>";
        echo "文件临时存储的位置: " . $_FILES["file"]["tmp_name"] . "<br>";
        
        // 判断当前目录下的 upload 目录是否存在该文件
        // 如果没有 upload 目录，你需要创建它，upload 目录权限为 777
        if (file_exists("uploadA/" . $_FILES["file"]["name"]))
        {
            echo $_FILES["file"]["name"] . " 文件已经存在。 ";
        }
        else
        {
            // 如果 upload 目录不存在该文件则将文件上传到 upload 目录下
            move_uploaded_file($_FILES["file"]["tmp_name"], "uploadA/" . $_FILES["file"]["name"]);
            echo "文件存储在: " . "uploadA/" . $_FILES["file"]["name"];
        }
    }
#header("content-type:text/html;charset=utf-8");
#exec("python3 preForA.py");
#header("Location: http://47.114.178.105/QualityCtrl/No1FileTool/resultA");

header("Location: http://47.114.178.105/QualityCtrl/No1FileTool/preForA.php");
}
else
{
    echo "上传失败，错误信息：";
    echo "非法的文件格式,！！注意！！仅能上传xlsx，不支持xls文件"."<br>";
#    echo "如需上传更大的文件，请联系管理员";
}

//header("content-type:text/html;charset=utf-8");
//exec('python3 preForA.py');
?>
