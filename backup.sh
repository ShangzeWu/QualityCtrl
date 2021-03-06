#! /bin/sh

find /var/www/html/QualityCtrl/No1FileTool/resultA/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/backupResultA \;

find /var/www/html/QualityCtrl/No1FileTool/resultB/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/backupResultB \;

find /var/www/html/QualityCtrl/No1FileTool/uploadA/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/backupUploadA \;

find /var/www/html/QualityCtrl/No1FileTool/uploadB/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/backupUploadB \;

find /var/www/html/QualityCtrl/No2FileTool/uploadC/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/backupUploadC \;

find /var/www/html/QualityCtrl/No2FileTool/resultTemp/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/resultTemp \;

find /var/www/html/QualityCtrl/No2FileTool/resultFinal/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/resultFinal \;

find /var/www/html/QualityCtrl/No2FileTool/resultC/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/backupResultC \;

find /var/www/html/QualityCtrl/No3FileTool/MixABD/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/MixABD \;

find /var/www/html/QualityCtrl/No3FileTool/uploadD/*  -name "*.xlsx*" -exec mv {} /var/www/html/backupFiles/uploadD \;

find /var/www/html/QualityCtrl/No1FileTool/MixAB/*  -name "*.xlsx*" -exec rm -f {} \;

find /var/www/html/QualityCtrl/No1FileTool/MixAB_B/*  -name "*.xlsx*" -exec rm -f {} \;

find /var/www/html/QualityCtrl/No1FileTool/MixAB_A/*  -name "*.xlsx*" -exec rm -f {} \;
