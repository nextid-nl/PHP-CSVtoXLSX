<?
@apache_setenv('no-gzip', 1);

include("xlsx.php");

$filetplpath = "test.xlsx";
$fileexportname = "testexport.xlsx";

$tmppath=tempnam('php://temp', 'xl_');
copy($filetplpath,$tmppath);

$xl=new xlsx($tmppath);
$xl->importcsv("csvdata","test.csv",",",'"','A','2');
$xl->refreshPivotsOnOpen ();
$xl->close();

header('Content-Disposition: attachment;filename="'.$fileexportname.'"');
ob_clean();   
readfile($tmppath);
exit;
?>
