<?
@apache_setenv('no-gzip', 1);

include("xlsx.php");

$filetpl = "test.xlsx";
$file = "testexport.xlsx";

$path=tempnam('php://temp', 'php');
copy($filetpl,$path);

$xl=new xlsx($path);
$xl->importcsv("csvdata","test.csv",",",'"','A','2');
$xl->refreshPivotsOnOpen ();
$xl->close();

header('Content-Disposition: attachment;filename="'.$file.'"');
ob_clean();   
readfile($path);
exit;
?>