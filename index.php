<?
@apache_setenv('no-gzip', 1);
include("xlsx.php");

/*settings*/
$filetplpath = "test.xlsx";
$fileexportname = "testexport.xlsx";
$csvpath1 = "test.csv";
$csvtosheet = "csvdata";
$csvtocol = "A";
$csvtorow = "2";
$csvdelimiter = ",";
$csvenclosure = '"';

/*create temp xlsx file*/
$tmppath=tempnam('php://temp', 'xl_');
copy($filetplpath,$tmppath);

/*open tempfile, import the csv file, and set pivots to refresh*/
$xl=new xlsx($tmppath);
$xl->importcsv($csvtosheet,$csvpath1,$csvdelimiter,$csvenclosure,$csvtocol,$csvtorow);
$xl->refreshPivotsOnOpen ();
$xl->close();

/*output tempfile*/
header('Content-Disposition: attachment;filename="'.$fileexportname.'"');
ob_clean();   
readfile($tmppath);
exit;
?>
