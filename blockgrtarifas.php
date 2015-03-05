<?php 

class blockgrtarifas extends Module {
	private $sheets;
	private $hyperlinks;
	private $package;
	private $sharedstrings;
	// esquema 
	private $SCHEMA_OFFICEDOCUMENT  =  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
	private $SCHEMA_RELATIONSHIP  =  'http://schemas.openxmlformats.org/package/2006/relationships';
	private $SCHEMA_SHAREDSTRINGS =  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
	private $SCHEMA_WORKSHEETRELATION =  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

	function __construct()
	{
        $this->name = "blockgrtarifas";
        $this->tab = 'blockgrtarifas';
        $this->version = '0.9.212';
        parent::__construct();
        $this->displayName = $this->l('Mantenimiento Tarifas');
        $this->description = $this->l('Importación de tarifas mediante Plantilla Excel 2007');
	}
	
	function install()
	{
            mkdir("../upload/tarifas/", 0777);
            mkdir("../upload/tarifas/src/", 0777);
       if (!parent::install()
           OR !$this->registerHook('rightColumn')
           OR !$this->registerHook('leftColumn'))
                return false;
        return true;	
	}
	
	public function hookLeftColumn($params)	
	{
		echo "Hello World!";
	}
	
    public function hookRightColumn($params)
    {
        return $this->hookLeftColumn($params);
    }
	
	function uninstal()
	{
                $this->rrmdir("../upload/tarifas/");
		if (!parent::uninstall())
			return false;
		return true;
	}
        
     public function displayForm()
	{
         if(isset($_REQUEST['ffd'])){
             $this->deleteFileOnServer($_REQUEST['ffd']);
             $output="Eliminado Archivo ".$_REQUEST['ffd'];
         }
		return '
		     <div style="border-right:10px;margin:0 auto;width:400px;height:100px;text-align:center;">
          <form  method="post" enctype="multipart/form-data">
         <input type="file" name="file1" value="Buscar archivo..."/><br/>
          <p style="font-size:15px;text-align: center;">

              
            <input type="submit" value="Actualizar tarifas"  name="submitSpecials"/><br/>
          </form>
      </div>
<div style="border-right:10px;margin:0 auto;width:400px;height:100px;text-align:center;">
          <form  method="post" >
        Generar un archivo con los datos de su tienda.<br/>
          <p style="font-size:15px;text-align: center;">

              
            <input type="submit" value="Generar Archivo"  name="submitFiles"/><br/>
          </form>
      </div><div style="border-right:10px;margin:0 auto;width:400px;height:100px;text-align:center;">'.$this->comprobarArchivo().'</div>';
	}
        
        	public function getContent()
	{
		$output = '<h2>'.$this->displayName.'</h2>';
		if (Tools::isSubmit('submitSpecials'))
		{
			
                    if ($_FILES["file1"]["error"] > 0)
                          {
                          $output.="<span style='color:red;'>yeeeeee !!!!!!  Error: " . $_FILES["file1"]["error"] ."</span>";
                          }
                        else
                          {
                          $random_digit=rand(0000,9999);
                          $nomArxiu=$random_digit.date('Y-m-d ').$_FILES["file1"]["name"];  
                          $output.="Pujat : " . $_FILES["file1"]["name"] . "\n";
                          $output.="Tipus arxiu : " . $_FILES["file1"]["type"] . "\n";
                          $output.="Mida : " . ($_FILES["file1"]["size"] / 1024) . " Kb \n";
                          $output.="Carpeta Temporal " . $_FILES["file1"]["tmp_name"]. " \n";
                          $output.="Guardat a/com : src/".$nomArxiu. " \n";
                          //echo($output);
                          copy($_FILES["file1"]["tmp_name"], '../upload/'.$nomArxiu);
                          
                          try {
                               $xlsx = $this->startParseFile('../upload/'.$nomArxiu);
                               $output.=$this->extreureTarifes($xlsx);
                            } catch (Exception $e) {
                                $output.='ERROR!!! ====> Excepción capturada: '.$e->getMessage()."\n";
                            }
                        }
                }
                
                if (Tools::isSubmit('submitFiles'))
		{
			
                    $this->createExcel();
                }
		return $output.$this->displayForm();
	}
        
    function extreureTarifes($xlsx){
        $num=0;
        foreach( $this->rows(1) as $k => $r) {
                        if(($r[0]!='')&&($r[0]!=' ')&&($r[0]!='id_product')){
                            $this->prestaTarifas($r[0],$r[5],$r[6]);
                            $num++;
                        }
            }

            return "\n Modificats : ".$num." registres \n";
    }

        function rrmdir($dir) {
           if (is_dir($dir)) {
             $objects = scandir($dir);
             foreach ($objects as $object) {
               if ($object != "." && $object != "..") {
                 if (filetype($dir."/".$object) == "dir") rrmdir($dir."/".$object); else unlink($dir."/".$object);
               }
             }
             reset($objects);
             rmdir($dir);
           }
 }
    
    
    function prestaTarifas($id,$preuCompra,$preuVenta){
        $conn =mysql_connect(_DB_SERVER_, _DB_USER_, _DB_PASSWD_);
        mysql_select_db(_DB_NAME_, $conn);
        //Consulta para actulizar datos en Prestashop
        mysql_query("UPDATE ps_product SET price = $preuVenta, wholesale_price = $preuCompra WHERE id_product = $id");
        mysql_close($conn);
    }
    
    function startParseFile( $filename ) {
		$this->_unzip( $filename );
		$this->_parse();
	}
	function sheets() {
		return $this->sheets;
	}
	function sheetsCount() {
		return count($this->sheets);
	}
	function worksheet( $worksheet_id ) {
            //echo "numero de worksheet"+$worksheet_id;
		if ( isset( $this->sheets[ $worksheet_id ] ) ) {
			$ws = $this->sheets[ $worksheet_id ];

			if (isset($ws->hyperlinks)) {
				$this->hyperlinks = array();
				foreach( $ws->hyperlinks->hyperlink as $hyperlink ) {
					$this->hyperlinks[ (string) $hyperlink['ref'] ] = (string) $hyperlink['display'];
				}
			}

			return $ws;
		} else
			throw new Exception('Worksheet '.$worksheet_id.' not found.');
	}
	function dimension( $worksheet_id = 1 ) {
		$ws = $this->worksheet($worksheet_id);
		$ref = (string) $ws->dimension['ref'];
		$d = explode(':', $ref);
		$index = $this->_columnIndex( $d[1] );
		return array( $index[0]+1, $index[1]+1);
	}
	// sheets numeration: 1,2,3....
	function rows( $worksheet_id ) {

		$ws = $this->worksheet( $worksheet_id);

		$rows = array();
		$curR = 0;

		foreach ($ws->sheetData->row as $row) {

			foreach ($row->c as $c) {
				list($curC,) = $this->_columnIndex((string) $c['r']);
				$rows[ $curR ][ $curC ] = $this->value($c);
			}

			$curR++;
		}
		return $rows;
	}
	function rowsEx( $worksheet_id  ) {
		$rows = array();
		$curR = 0;
		if (($ws = $this->worksheet( $worksheet_id)) === false)
			return false;
		foreach ($ws->sheetData->row as $row) {

			foreach ($row->c as $c) {
				list($curC,) = $this->_columnIndex((string) $c['r']);
				$rows[ $curR ][ $curC ] = array(
					'name' => (string) $c['r'],
					'value' => $this->value($c),
					'href' => $this->href( $c ),
				);
			}
			$curR++;
		}
		return $rows;

	}
	// thx Gonzo
	function _columnIndex( $cell = 'A1' ) {

		if (preg_match("/([A-Z]+)(\d+)/", $cell, $matches)) {

			$col = $matches[1];
			$row = $matches[2];

			$colLen = strlen($col);
			$index = 0;

			for ($i = $colLen-1; $i >= 0; $i--)
				$index += (ord($col{$i}) - 64) * pow(26, $colLen-$i-1);

			return array($index-1, $row-1);
		} else
			throw new Exception("Invalid cell index.");
	}
	function value( $cell ) {
		// Determine data type
		$dataType = (string)$cell["t"];
		switch ($dataType) {
			case "s":
				// Value is a shared string
				if ((string)$cell->v != '') {
					$value = $this->sharedstrings[intval($cell->v)];
				} else {
					$value = '';
				}

				break;

			case "b":
				// Value is boolean
				$value = (string)$cell->v;
				if ($value == '0') {
					$value = false;
				} else if ($value == '1') {
					$value = true;
				} else {
					$value = (bool)$cell->v;
				}

				break;

			case "inlineStr":
				// Value is rich text inline
				$value = $this->_parseRichText($cell->is);

				break;

			case "e":
				// Value is an error message
				if ((string)$cell->v != '') {
					$value = (string)$cell->v;
				} else {
					$value = '';
				}

				break;

			default:
				// Value is a string
				$value = (string)$cell->v;

				// Check for numeric values
				if (is_numeric($value) && $dataType != 's') {
					if ($value == (int)$value) $value = (int)$value;
					elseif ($value == (float)$value) $value = (float)$value;
					elseif ($value == (double)$value) $value = (double)$value;
				}
		}
		return $value;
	}
	function href( $cell ) {
		return isset( $this->hyperlinks[ (string) $cell['r'] ] ) ? $this->hyperlinks[ (string) $cell['r'] ] : '';
	}
	function _unzip( $filename ) {
		// Clear current file
		$this->datasec = array();

		// Package information
		$this->package = array(
			'filename' => $filename,
			'mtime' => filemtime( $filename ),
			'size' => filesize( $filename ),
			'comment' => '',
			'entries' => array()
		);
        // Read file
		$oF = fopen($filename, 'rb');
		$vZ = fread($oF, $this->package['size']);
		fclose($oF);
		// Cut end of central directory
		$aE = explode("\x50\x4b\x05\x06", $vZ);

		// Normal way
		$aP = unpack('x16/v1CL', $aE[1]);
		$this->package['comment'] = substr($aE[1], 18, $aP['CL']);

		// Translates end of line from other operating systems
		$this->package['comment'] = strtr($this->package['comment'], array("\r\n" => "\n", "\r" => "\n"));

		// Cut the entries from the central directory
		$aE = explode("\x50\x4b\x01\x02", $vZ);
		// Explode to each part
		$aE = explode("\x50\x4b\x03\x04", $aE[0]);
		// Shift out spanning signature or empty entry
		array_shift($aE);

		// Loop through the entries
		foreach ($aE as $vZ) {
			$aI = array();
			$aI['E']  = 0;
			$aI['EM'] = '';
			// Retrieving local file header information
//			$aP = unpack('v1VN/v1GPF/v1CM/v1FT/v1FD/V1CRC/V1CS/V1UCS/v1FNL', $vZ);
			$aP = unpack('v1VN/v1GPF/v1CM/v1FT/v1FD/V1CRC/V1CS/V1UCS/v1FNL/v1EFL', $vZ);
			// Check if data is encrypted
//			$bE = ($aP['GPF'] && 0x0001) ? TRUE : FALSE;
			$bE = false;
			$nF = $aP['FNL'];
			$mF = $aP['EFL'];

			// Special case : value block after the compressed data
			if ($aP['GPF'] & 0x0008) {
				$aP1 = unpack('V1CRC/V1CS/V1UCS', substr($vZ, -12));

				$aP['CRC'] = $aP1['CRC'];
				$aP['CS']  = $aP1['CS'];
				$aP['UCS'] = $aP1['UCS'];

				$vZ = substr($vZ, 0, -12);
			}

			// Getting stored filename
			$aI['N'] = substr($vZ, 26, $nF);
			if (substr($aI['N'], -1) == '/') {
				// is a directory entry - will be skipped
				continue;
			}

			// Truncate full filename in path and filename
			$aI['P'] = dirname($aI['N']);
			$aI['P'] = $aI['P'] == '.' ? '' : $aI['P'];
			$aI['N'] = basename($aI['N']);

			$vZ = substr($vZ, 26 + $nF + $mF);

			if (strlen($vZ) != $aP['CS']) {
			  $aI['E']  = 1;
			  $aI['EM'] = 'Compressed size is not equal with the value in header information.';
			} else {
				if ($bE) {
					$aI['E']  = 5;
					$aI['EM'] = 'File is encrypted, which is not supported from this class.';
				} else {
					switch($aP['CM']) {
						case 0: // Stored
							// Here is nothing to do, the file ist flat.
							break;
						case 8: // Deflated
							$vZ = gzinflate($vZ);
							break;
						case 12: // BZIP2
							if (! extension_loaded('bz2')) {
								if (strtoupper(substr(PHP_OS, 0, 3)) == 'WIN') {
								  @dl('php_bz2.dll');
								} else {
								  @dl('bz2.so');
								}
							}
							if (extension_loaded('bz2')) {
								$vZ = bzdecompress($vZ);
							} else {
								$aI['E']  = 7;
								$aI['EM'] = "PHP BZIP2 extension not available.";
							}
							break;
						default:
						  $aI['E']  = 6;
						  $aI['EM'] = "De-/Compression method {$aP['CM']} is not supported.";
					}
					if (! $aI['E']) {
						if ($vZ === FALSE) {
							$aI['E']  = 2;
							$aI['EM'] = 'Decompression of data failed.';
						} else {
							if (strlen($vZ) != $aP['UCS']) {
								$aI['E']  = 3;
								$aI['EM'] = 'Uncompressed size is not equal with the value in header information.';
							} else {
								if (crc32($vZ) != $aP['CRC']) {
									$aI['E']  = 4;
									$aI['EM'] = 'CRC32 checksum is not equal with the value in header information.';
								}
							}
						}
					}
				}
			}

			$aI['D'] = $vZ;

			// DOS to UNIX timestamp
			$aI['T'] = mktime(($aP['FT']  & 0xf800) >> 11,
							  ($aP['FT']  & 0x07e0) >>  5,
							  ($aP['FT']  & 0x001f) <<  1,
							  ($aP['FD']  & 0x01e0) >>  5,
							  ($aP['FD']  & 0x001f),
							  (($aP['FD'] & 0xfe00) >>  9) + 1980);

			//$this->Entries[] = &new SimpleUnzipEntry($aI);
			$this->package['entries'][] = array(
				'data' => $aI['D'],
				'error' => $aI['E'],
				'error_msg' => $aI['EM'],
				'name' => $aI['N'],
				'path' => $aI['P'],
				'time' => $aI['T']
			);

		} // end for each entries
	}
	function getPackage() {
		return $this->package;
	}
	function getEntryData( $name ) {
		$dir = dirname( $name );
		$name = basename( $name );
		foreach( $this->package['entries'] as $entry)
			if ( $entry['path'] == $dir && $entry['name'] == $name)
				return $entry['data'];
	}
	function unixstamp( $excelDateTime ) {
		$d = floor( $excelDateTime ); // seconds since 1900
		$t = $excelDateTime - $d;
		return ($d > 0) ? ( $d - 25569 ) * 86400 + $t * 86400 : $t * 86400;
	}
	function _parse() {
		// Document data holders
		$this->sharedstrings = array();
		$this->sheets = array();

		// Read relations and search for officeDocument

		$relations = simplexml_load_string( $this->getEntryData("_rels/.rels") );
		foreach ($relations->Relationship as $rel) {
			if ($rel["Type"] == $this->SCHEMA_OFFICEDOCUMENT) {
				// Found office document! Read relations for workbook...
				$workbookRelations = simplexml_load_string($this->getEntryData( dirname($rel["Target"]) . "/_rels/" . basename($rel["Target"]) . ".rels") );
				$workbookRelations->registerXPathNamespace("rel", $this->SCHEMA_RELATIONSHIP);

				// Read shared strings
				$sharedStringsPath = $workbookRelations->xpath("rel:Relationship[@Type='" . $this->SCHEMA_SHAREDSTRINGS . "']");
				$sharedStringsPath = (string)$sharedStringsPath[0]['Target'];
				$xmlStrings = simplexml_load_string($this->getEntryData( dirname($rel["Target"]) . "/" . $sharedStringsPath) );
				if (isset($xmlStrings) && isset($xmlStrings->si)) {
					foreach ($xmlStrings->si as $val) {
						if (isset($val->t)) {
							$this->sharedstrings[] = (string)$val->t;
						} elseif (isset($val->r)) {
							$this->sharedstrings[] = $this->_parseRichText($val);
						}
					}
				}

				// Loop relations for workbook and extract sheets...
				foreach ($workbookRelations->Relationship as $workbookRelation) {
					if ($workbookRelation["Type"] == $this->SCHEMA_WORKSHEETRELATION) {
						$this->sheets[ str_replace( 'rId', '', (string) $workbookRelation["Id"]) ] =
							simplexml_load_string( $this->getEntryData( dirname($rel["Target"]) . "/" . dirname($workbookRelation["Target"]) . "/" . basename($workbookRelation["Target"])) );
					}
				}

				break;
			}
		}

		// Sort sheets
		ksort($this->sheets);
	}
    private function _parseRichText($is = null) {
        $value = array();

        if (isset($is->t)) {
            $value[] = (string)$is->t;
        } else {
            foreach ($is->r as $run) {
                $value[] = (string)$run->t;
            }
        }

        return implode(' ', $value);
    }
	function getWorksheetName($dimId = 0){
		$worksheetName = array();
		$xmlWorkBook = simplexml_load_string( $this->getEntryData("xl/workbook.xml") );
		if($dimId==0){
			foreach ($xmlWorkBook->sheets->sheet as $sheetName) {
				$worksheetName[] = $sheetName['name'];
			}
		}else{
			$worksheetName[] = $xmlWorkBook->sheets->sheet[$dimId-1]->attributes()->name;
		}
		return $worksheetName;
	}
     // Get Shop Product Taxes & Specs.   
    function getTarifas(){
        $conn =mysql_connect(_DB_SERVER_, _DB_USER_, _DB_PASSWD_);
        mysql_select_db(_DB_NAME_, $conn);
        //Consulta para actulizar datos en Prestashop
        $result=mysql_query("SELECT DISTINCT(a.id_product),
            (SELECT b.name FROM ps_product_lang b WHERE b.id_product=a.id_product LIMIT 1) AS name, 
            a.reference, 
            (SELECT b.description FROM ps_product_lang b WHERE b.id_product=a.id_product LIMIT 1) AS description, 
            (SELECT b.description_short FROM ps_product_lang b WHERE b.id_product=a.id_product LIMIT 1) AS description_short, 
            a.price, 
            a.wholesale_price, 
            a.date_add  
            FROM ps_product a") or die('\n Error Consulta: ' . mysql_error().' \n ');
        mysql_close($conn);
        return $result;
    }
        
        //EXCEL WRITTER
     function createExcel(){   
        /** Error reporting */
        error_reporting(E_ALL);

        /** Include path **/
        ini_set('include_path', ini_get('include_path').';../Classes/');

        /** PHPExcel */
        include 'PHPExcel.php';

        /** PHPExcel_Writer_Excel2007 */
        include 'PHPExcel/Writer/Excel2007.php';

        // Create new PHPExcel object
        echo date('H:i:s') . " Create new PHPExcel object\n";
        $objPHPExcel = new PHPExcel();

        // Set properties
        echo date('H:i:s') . " Set properties\n";
        $objPHPExcel->getProperties()->setCreator("Tienda Online");
        $objPHPExcel->getProperties()->setLastModifiedBy("Tienda Online");
        $objPHPExcel->getProperties()->setTitle("Archivo de Tarifas Office 2007 XLSX");
        $objPHPExcel->getProperties()->setSubject("Archivo de Tarifas Office 2007 XLSX");
        $objPHPExcel->getProperties()->setDescription("Archivo generado para el mantenimiento de Productos de la Tienda Online");

        $result=$this->getTarifas();
        $act_row=1;
        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->getActiveSheet()->SetCellValue('A'.$act_row, 'id_product');
            $objPHPExcel->getActiveSheet()->SetCellValue('B'.$act_row, 'name');
            $objPHPExcel->getActiveSheet()->SetCellValue('C'.$act_row, 'reference');
            $objPHPExcel->getActiveSheet()->SetCellValue('D'.$act_row, 'description');
            $objPHPExcel->getActiveSheet()->SetCellValue('E'.$act_row, 'description_short');
            $objPHPExcel->getActiveSheet()->SetCellValue('F'.$act_row, 'price');
            $objPHPExcel->getActiveSheet()->SetCellValue('G'.$act_row, 'wholesale_price');
            $objPHPExcel->getActiveSheet()->SetCellValue('H'.$act_row, 'date_add');
            $act_row++;
        while($row = mysql_fetch_array($result)){
            $objPHPExcel->getActiveSheet()->SetCellValue('A'.$act_row, $row['id_product']);
            $objPHPExcel->getActiveSheet()->SetCellValue('B'.$act_row, $row['name']);
            $objPHPExcel->getActiveSheet()->SetCellValue('C'.$act_row, $row['reference']);
            $objPHPExcel->getActiveSheet()->SetCellValue('D'.$act_row, $row['description']);
            $objPHPExcel->getActiveSheet()->SetCellValue('E'.$act_row, $row['description_short']);
            $objPHPExcel->getActiveSheet()->SetCellValue('F'.$act_row, $row['price']);
            $objPHPExcel->getActiveSheet()->SetCellValue('G'.$act_row, $row['wholesale_price']);
            $objPHPExcel->getActiveSheet()->SetCellValue('H'.$act_row, $row['date_add']);
            $act_row++;
        }

        // Rename sheet
        $objPHPExcel->getActiveSheet()->setTitle('Tarifas ');


        // Save Excel 2007 file
        //echo date('H/i/s') . " Write to Excel2007 format\n";
        $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
        //alimentamos el generador de aleatorios
        srand (time());
        //generamos un número aleatorio
        $numero_aleatorio = rand(1,100000); 
        echo(__FILE__);
        $objWriter->save(str_replace('.php', '.xlsx', __FILE__));
        rename(str_replace('.php','.xlsx', __FILE__), "../upload/tarifas/src/tarifas".$numero_aleatorio.".xlsx");
        //'../upload/tarifas/tarifas'.$numero_aleatorio.'.xlsx');//s

        
        
        

    
     }  
     //FORÇAR DESCARGA PER A UN ARXIU (no utiliztat)
     function forzarDescarga(){
        $enlace = '../modules/blockgrtarifas/blockgrtarifas.xlsx';
        header('Content-Description: File Transfer');
        header('Content-Type: application/octet-stream');
        header('Content-Disposition: attachment; filename='.basename($enlace));
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Content-Length: ' . filesize($enlace));
        ob_clean();
        flush();
        readfile($enlace);
     }
     //COMPROVEM ELS ARXIUS EXISTENTS EN EL SERVIDOR 
     function comprobarArchivo(){
         $ruta= '../upload/tarifas/src/';
         $salida='';
         $count=0;
         $count2=0;
       // abrir un directorio y listarlo recursivo
       if (is_dir($ruta)) {
          if ($dh = opendir($ruta)) {
             while (($file = readdir($dh)) !== false) {
                //esta línea la utilizaríamos si queremos listar todo lo que hay en el directorio
                //mostraría tanto archivos como directorios
                //echo "<br>Nombre de archivo: $file : Es un: " . filetype($ruta . $file);
                 //echo "<p> resltat ->> ".substr($file,-4)."</p><br>";
                 
                 $count++;
                if (substr($file,-4)== 'xlsx' ){
                    $count++;
                   //solo si el archivo es un directorio, distinto que "." y ".."... i te extenció xlsx 
                   
                    $salida.="<span ><a href='../upload/tarifas/src/$file' alt='file'>$file</a></span> <span><a href='index.php?tab=AdminModules&configure=blockgrtarifas&token=".$_REQUEST['token']."&ffd=$file' alt='file'><img src='../modules/blockgrtarifas/img/delete.png' alt='delete'/></a></span></br>";
                    //echo "<br>Directorio: $ruta$file";
                }
             }
    
            if($count<=2){
                  $salida.="<br>No Existe ningun archivo en el servidor";  
            }
          closedir($dh);
          }
       }else
          $salida.="<br>No es ruta valida";
            
        
          return $salida;
     }
     
     function deleteFileOnServer($file){
         unlink('../upload/tarifas/src/'.$file);
     }
}
?>