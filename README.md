PHP CSVtoXLSX
============

class xlsx

simple class to load csvfile(s) data in xlsx worksheets and export to user:
    
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
    
    
cons
- you can not directly manipulate cells / rows / sheets / workbook
- there is no calculation engine

pros
- will not touch sheets other than the csv loaded sheet and sheets with references
- will not create a new archive but only updates the sheetdata in the existing archive
- workbook/sheet/row/cell settings are all kept during csv load
- references to overwritten cells are updated
- easy class for filling a prepared template with csv data, endless possibilities, example usage: 
    1 create a nice looking template with graphs,etc 
    2 make an empty extra sheet per existing sheet (the datasheet where the nice sheet will get it's data from), 
    3 make references on the existing sheet to the empty one (look at the empty one as if it is the csv) 
    4 Hide the empty sheet
    5 use this class to fill the hidden empty sheet with csv data.
    6 nice sheet will be filled by the references made in step 3
    
    
so what about extra functionality? 
... you can add methods to do so 

- cellmanipulation could be easy with xpath "//n:c[@r=A1]" , where A1 is the cell-reference (colletter-rownumber) 
- rowmanipulation could be easy with xpath "//n:row[@r=1]" , where 1 is the row-reference (rownumber) 

set-cell-value examplecode:
    
        $sheet=$this->get_sheet("sheet1name");
        $sheetxml=$this->zip->getFromName("xl/".$sheet["target"]);
        $dom=new DOMDocument();
        $dom->loadXML($sheetxml);
        $xp=new DOMXpath($dom);
        $cell=$xp->query("//n:c[@r='A1']")->item(0);
        if(!$cell) {//sheets not always have the cell
            $cell=$dom->createElement('c.....
            ...
        }
        setInnerXML($cell,($numeric?"<v>".$value."</v>":"<is><t>".$value."</t></is>"));
        $this->_create_zipfile("xl/".$sheet['target'], $dom->saveXML());
    
keep in mind:
to build excel workbooks, manipulating styles, conditional formatting, calculations etc... 
all harder to implement and it will increase this class code complexity. 
Use PHPExcel for that!
    
    
Norbert Peters (norbert@nextid.nl)
