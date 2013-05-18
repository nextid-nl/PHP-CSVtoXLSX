<?
/*
class xlsx
simple class to load csvfile data in xlsx worksheets
    
    @apache_setenv('no-gzip', 1);
    copy("[xlsx_templatefilepath]","[xlsx_newfilepath]");
    $xl=new xlsx("[xlsx_newfilepath]");
    $xl->importcsv("[sheet1name]","[csv1_filepath]");
    $xl->importcsv("[sheet2name]","[csv2_filepath]");
    $xl->refreshPivotsOnOpen ();
    $xl->close();
    header('Content-Disposition: attachment;filename="[xlsx_newfilename]"');
    ob_clean();   
    readfile("[xlsx_newfilepath]");
    
Norbert Peters 
norbert@nextid.nl
*/
class xlsx {
    private $worksheets=array();/*worsheets array(sheet hash(name => [sheet name], target => [sheet file ref in zip])[,sheet hash])*/
    private $zip;
    private $file;
    
    function __construct($spath) {
        $this->zip=new ZipArchive();
        $this->open($spath);
    }
    /*open a xlsx archive with php_zip and get the sheet settings, 
    open method is public to reset object and load a new archive*/
    public function open($spath) {
        $this->file=$spath;
        $this->_open_zip();
        $this->worksheets=$this->get_sheetsettings();
    }
    private function get_booksettings_dom($xlrels) {
        $xlrels=simplexml_load_string($this->zip->getFromName("_rels/.rels")); //workbook-target is found in here
        /*get workbook.xml target for sheetsettings and pivotcachesettings*/
        $target="";
        foreach($xlrels->Relationship as $rel) {if($rel["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument") $target=$rel["Target"];}
        if(!$target) die("<b>Error (open):</b> xlsx corrupt (workbook.xml is missing)");
        /*read workbook.xml*/
        $wbxml=$this->zip->getFromName($target);
        if(!$wbxml) die("<b>Error (open):</b> xlsx corrupt (workbookfile '".$target."' not found/empty)");
        $xmlbook=new DOMDocument();
        $xmlbook->loadXML($wbxml);
        return $xmlbook;
    }
    private function get_sheetsettings() {
        $worksheets=array();
        $wbrels=simplexml_load_string($this->zip->getFromName("xl/_rels/workbook.xml.rels")); //sheet-targets are found in here
        $xmlbook=$this->get_booksettings_dom($xlrels);
        /*get the sheets and store settings in this.worksheets array (extra settings are not stored (only the required for importcsv))*/
        $sheets=$xmlbook->getElementsByTagName("sheet");
        if(!$sheets||!$sheets->length) die("<b>Error (open):</b> xlsx corrupt (sheet links in workbook.xml are missing)");
        foreach($sheets as $elsheet) {
            $sheet=array();
            $sheet["name"]=(string) $elsheet->getAttribute("name");
            /*get target of the sheet from the workbook rels*/
            foreach($wbrels->Relationship as $xlrel) { 
                if($xlrel["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" && $xlrel["Id"]==$elsheet->getAttribute("r:id")) 
                    $sheet["target"]=(string) $xlrel["Target"];
            }
            $worksheets[]=$sheet;
        }
        return $worksheets;
    }
    public function importcsv($sheetname,$csvstring_or_path,$csvdelimiter="\t",$csvenclosure='"',$col_ref="A",$row_ref=1,$skip_empty_cells=true) {
        $sheet=$this->get_sheet($sheetname);
        $oldsheetxml=$this->zip->getFromName("xl/".$sheet["target"]);
        if(!$oldsheetxml) die("<b>Error (importcsv):</b> xlsx corrupt (sheetfile '"."xl/".$sheet["target"]."' not found/empty)");
        /*
        manipulation via DOMDocument 
            old sheet is loaded as xml DOMDocument, 
            xpath is used to get cells and rows by OpenXML attribute r (reference))
            not existing cells/rows that does exist in the csv will be added to sheetxml
        */
        $dom=new DOMDocument();
        $dom->loadXML($oldsheetxml);
        $xp=new DOMXpath($dom);
        $xp->registerNamespace('n', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        
        /*if param is a textstring treat as delimited / comma seperated values (csv string) to do so write to temp file which will be readed with fgetcsv (compatible with versions before PHP 5.3.0 as str_getcsv doesn't)*/
        if(!is_file($csvstring_or_path)) {
            $handle=tmpfile();
            fwrite($handle, $csvstring_or_path); 
            fseek($handle, 0); 
        }
        /*if param is a path to file treat as csv file*/
        else $handle=fopen($csvstring_or_path, "r");
        
        if($handle !== FALSE) {//open csv else do nothing
            $rownumber=$row_ref;
            while(($data=fgetcsv($handle, 1000, $csvdelimiter)) !== FALSE) {//read csv per line
                $colchar=$col_ref;
                $prevcolref="";//reference where to put a new colnode if not exists (after prevcolref)
                foreach($data as $v) {
                    if($v==""&&$skip_empty_cells) {//skip column, do nothing if column in csv has nothing
                        $colchar++;
                        continue;
                    }
                    $colref=$colchar.$rownumber;//current cell to write
                    $numeric=is_numeric($v);
                    $valnode=($numeric?"<v>".$v."</v>":"<is><t>".$v."</t></is>");//how to write in excel (strings via inline richtext, then -> no need to manipulate sharedstrings.xml)
                    $cell=$xp->query("//n:c[@r='".$colref."']")->item(0);
                    if(!$cell) {//cell not found then create cell
                        $row=$xp->query("//n:row[@r='".$rownumber."']")->item(0);
                        if($row) $cell=$dom->createElement('c');
                        else {//row not found then create row
                            $sheetdata=$dom->getElementsByTagName("sheetData")->item(0);
                            $row=$dom->createElement('row');
                            $row->setAttribute('r',$rownumber);
                            /*insert row in right order to prevent excel from corruptionwarning*/
                            $nextrow=$xp->query("//n:row[@r='".($rownumber + 1)."']")->item(0);
                            if($nextrow) $sheetdata->insertBefore($row,$nextrow);
                            /*not found then append*/
                            else $sheetdata->appendChild($row);
                            $cell=$dom->createElement('c');
                        }
                        $cell->setAttribute('r',$colref);
                        /*insert cell in right order to prevent excel from corruptionwarning*/
                        $nextcell=$xp->query("//n:c[@r='".$prevcolref."']")->item(0); //pref
                        if($nextcell) $nextcell=$nextcell->nextSibling; //pref->nextSibling (to call insertBefore to)
                        if($nextcell) $row->insertBefore($cell,$nextcell);
                        /*not found then append*/
                        else $row->appendChild($cell);
                    }
                    setInnerXML($cell,$valnode);//set new cellvalue in DOMDocument
                    if(!$numeric) $cell->setAttribute("t","inlineStr");//strings via inline richtext
                    else $cell->removeAttribute("t");//numbers has no type attribute
                    $prevcolref=$colref; //prevcolref to insert cols after
                    $colchar++;//++ works for chars: A++ = B
                }
                $rownumber++;
            }
            $this->_create_zipfile("xl/".$sheet['target'], $dom->saveXML()); //overwrites old sheet with new sheet generated via xml DOMDocument
            $this->_delete_zipfile("xl/calcChain.xml"); //if not removing this file, excel displays message on open of created file to repair (if the csv data is written over cells that used to have formula's/references). When deleted, on open, excel will generate a new calcChain without warning
            $this->on_all_sheets_reset_formulareferences_to($sheetname);//remove the hard-writed values of cells with references to this sheet, excel will refresh them on open without any warning
            fclose($handle); 
            return true;
        }
        return false;
    }
    /*deletes v nodes (cached values) of cells that have formula's with a reference to $sheetname*/
    public function on_all_sheets_reset_formulareferences_to($sheetname) {
        foreach($this->worksheets as $sheet) {
            $dom=new DOMDocument();
            $dom->loadXML($this->zip->getFromName("xl/".$sheet["target"]));
            $fmls=$dom->getElementsByTagName('f');
            $dirtysheet=false;//only update sheets that have references
            foreach($fmls as $fmlo) {
                if(strpos($fmlo->nodeValue,"'")===false&&$sheet['name']!=$sheetname) continue; /*if no quote in formula and current loop sheet is not $sheetname then continue, only formulas without quote on $sheetname should be dirty (delete <v>)*/
                if(strpos($fmlo->nodeValue,"'".$sheetname."'")===false&&strpos($fmlo->nodeValue,"'")!==false) continue; /*if not referenced to sheetname then continue, on all other where referenced to '$sheetname' -> make dirty (delete <v>)*/
                $valnode=$fmlo->parentNode->getElementsByTagName('v')->item(0);
                $fmlo->parentNode->removeChild($valnode);
                $dirtysheet=true;
            }
            if($dirtysheet) {
                $this->_delete_zipfile("xl/".$sheet['target']);
                $this->_create_zipfile("xl/".$sheet['target'], $dom->saveXML());
            }
        }
    }
    /*set pivots to refresh data on open of the workbook*/
    public function refreshPivotsOnOpen () {
        $wbrels=simplexml_load_string($this->zip->getFromName("xl/_rels/workbook.xml.rels")); //pivotcache-targets are found in here
        $xmlbook=$this->get_booksettings_dom($xlrels);
        $pivotcaches=$xmlbook->getElementsByTagName("pivotCache");
        foreach($pivotcaches as $pcache) {
            $pivotcache=array();
            $pivotcache["cId"]=(string) $pcache->getAttribute("cacheId");
            /*get target of the cachedef from the workbook rels*/
            foreach($wbrels->Relationship as $xlrel) { 
                if($xlrel["Type"]=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" && $xlrel["Id"]==$pcache->getAttribute("r:id")) {
                    $target=(string)$xlrel["Target"];
                    $cachefilexml=$this->zip->getFromName("xl/".$target);
                    $dom=new DOMDocument();
                    $dom->loadXML($cachefilexml);
                    $dom->getElementsByTagName('pivotCacheDefinition')->item(0)->setAttribute('refreshOnLoad','1');
                }
            }
        }
    }
    /*get a sheet from the worksheets array */
    public function get_sheet($name) {
        $i=$this->get_sheetindex($name);
        return $this->worksheets[$i];
    }
    /*find the index from a sheet in the worksheets array */
    public function get_sheetindex($sheetname) {
        for($i=0;$i<count($this->worksheets);$i++) {
            if($this->worksheets[$i]["name"]==$sheetname) return $i;
        }
        return -1;
    }
    /*public close alias to close zip / force save and exit*/
    public function close() {
        $this->_close_zip();
    }
    /*opens the xlsx (zip) archive*/
    private function _open_zip() {
        if($this->zip->open($this->file)!==true) die("<b>Error (_open_zip):</b> xlsx corrupt (unpacking error)");
    }
    /*closes the xlsx (zip) archive*/
    private function _close_zip() {
        if($this->zip->close()!==true) die("<b>Error (_close_zip):</b> xlsx corrupt (saving error)");
    }
    /*deletes files in the xlsx (zip) archive*/
    private function _delete_zipfile($target) {
        $this->zip->deleteName($target);
        $this->_close_zip(); //after delete need to close and open zip to save data and to get new zip-file-index (otherwise targets are not correctly queried)
        $this->_open_zip();
    }
    /*creates a textfile in the xlsx (zip) archive filled with given textstring*/
    private function _create_zipfile($target,$tpl) {
        $this->zip->addFromString($target,$tpl);
        $this->_close_zip();//after create need to close and open zip to save data and to get new zip-file-index (otherwise targets are not correctly queried)
        $this->_open_zip();
    }
}

/*load xml string in node*/
function setInnerXML($ele, $xml) { 
    /*auth: Neil C. Obremski*/
    $idom=new DOMDocument();
    foreach($ele->childNodes as $child) {$ele->removeChild($child);}
    if(!$idom->loadXML("<x>{$xml}</x>")) return false;
    $import=$ele->ownerDocument->importNode($idom->documentElement, true);
    $i=0;
    while($i < ($len=$import->childNodes->length)) {
        $ele->appendChild($import->childNodes->item($i));
        if($len == $import->childNodes->length) $i++;
    }
    return true;
}