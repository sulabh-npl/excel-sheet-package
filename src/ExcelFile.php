<?php

namespace Sulabh\ExcelFilePackage;

use ZipArchive;

class ExcelFile
{
    /**
     * Create excel file from data
     * 
     * @param array $data
     * @param int $chunk_size
     * @param array $headers
     * @param string $filename - complete path of the file without extension
     * @param string $row_formatter - function to format each row
     * @param int $total_data_count - total data count (optional)
     * 
     * @created Sulabh Nepal
     * 
     * @return void
     */
    public static function createExcelFile($data, $chunk_size, $headers, $filename, $row_formatter, $total_data_count = 0){

        try{

            // echo("Creating XLSX file.");

            $total_data_count = $total_data_count ?? $data->count();

            $basePath = dirname(__DIR__, 4); 

            $publicPath = $basePath . '/public';
            
            if(!file_exists($publicPath.'/exports')) {
                mkdir($publicPath.'/exports');
            }

            str_replace('.xlsx', '', $filename);

            $filename = $publicPath.'/exports/'.$filename;

            mkdir($filename);

            mkdir($filename.'/docProps');

            mkdir($filename.'/_rels');

            mkdir($filename.'/xl');

            mkdir($filename.'/xl/_rels');

            mkdir($filename.'/xl/theme');

            mkdir($filename.'/xl/worksheets');

            mkdir($filename.'/xl/worksheets/_rels');

            $chunk_count = 0;

            $n_chunks = ceil($total_data_count / $chunk_size);

            $content_type_xml ='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';

            $docProcs_app_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>'.$n_chunks.'</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="'.$n_chunks.'" baseType="lpstr">';

            $xl_rels_workbook_xml_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';

            $xl_workbook_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/><workbookPr codeName="ThisWorkbook"/><bookViews><workbookView activeTab="0" autoFilterDateGrouping="true" firstSheet="0" minimized="false" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="600" visibility="visible"/></bookViews><sheets>';
                
            for($i = 1; $i <= $n_chunks; $i++) {

                $xl_worksheets_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';
            
                file_put_contents($filename.'/xl/worksheets/_rels/sheet'.$i.'.xml.rels', $xl_worksheets_rels);

                $content_type_xml .= '<Override PartName="/xl/worksheets/sheet'.$i.'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';

                $docProcs_app_xml .= '<vt:lpstr>Sheet'.$i.'</vt:lpstr>';

                $xl_rels_workbook_xml_rels .= '<Relationship Id="rId'.($i+3).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'.$i.'.xml"/>';

                $xl_workbook_xml .= '<sheet name="Sheet'.$i.'" sheetId="'.$i.'" r:id="rId'.($i+3).'"/>';

            }
            
            $content_type_xml .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>';

            $docProcs_app_xml .= '</vt:vector></TitlesOfParts><Company></Company><Manager></Manager><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinkBase></HyperlinkBase><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>';

            $xl_rels_workbook_xml_rels .= '</Relationships>';

            $xl_workbook_xml .= '</sheets><definedNames/><calcPr calcId="999999" calcMode="auto" calcCompleted="0" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>';

            file_put_contents($filename.'/[Content_Types].xml', $content_type_xml);

            file_put_contents($filename.'/docProps/app.xml', $docProcs_app_xml);

            file_put_contents($filename.'/xl/workbook.xml', $xl_workbook_xml);

            file_put_contents($filename.'/_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');

            file_put_contents($filename.'/xl/_rels/workbook.xml.rels', $xl_rels_workbook_xml_rels);

            // required to avoid memory exhaustion
            unset($content_type_xml, $docProcs_app_xml, $xl_rels_workbook_xml_rels, $xl_workbook_xml, $xl_worksheets_rels);

            $docProcs_core_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Sulabh Nepal</dc:creator><cp:lastModifiedBy>Sulabh Nepal</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">'.date('Y-m-d').'T'.date('H:i:s').'+00:00'.'</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">'.date('Y-m-d').'T'.date('H:i:s').'+00:00'.'</dcterms:modified><dc:title>EFI Export</dc:title><dc:description></dc:description><dc:subject></dc:subject><cp:keywords></cp:keywords><cp:category></cp:category></cp:coreProperties>';

            file_put_contents($filename.'/docProps/core.xml', $docProcs_core_xml);

            $xl_theme_theme1_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
        
            file_put_contents($filename.'/xl/theme/theme1.xml', $xl_theme_theme1_xml);

            $xl_styles_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="0"/><fonts count="1"><font><b val="0"/><i val="0"/><strike val="0"/><u val="none"/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf xfId="0" fontId="0" numFmtId="0" fillId="0" borderId="0" applyFont="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotTableStyle1"/></styleSheet>';

            file_put_contents($filename.'/xl/styles.xml', $xl_styles_xml);

            unset($xl_theme_theme1_xml, $docProcs_core_xml, $xl_styles_xml);

            $sharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

            $header_length = count($headers);

            $sharedStrings .= '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="'.(($total_data_count+1)*$header_length).'">';

            foreach($headers as $header) {
                $sharedStrings .= '<si><t>'.$header.'</t></si>';
            }

            touch($filename.'/xl/sharedStrings.xml');

            $shared_string_file = fopen($filename.'/xl/sharedStrings.xml', 'a');

            fwrite($shared_string_file, $sharedStrings);

            $shared_string_count = $header_length -1;

            $data->chunk($chunk_size, function($chunks) use ($filename, &$chunk_count, &$shared_string_file, &$shared_string_count, $header_length, $row_formatter, $chunk_size) {

                $chunk_count++;

                touch($filename.'/xl/worksheets/sheet'.$chunk_count.'.xml');

                $xml = fopen($filename.'/xl/worksheets/sheet'.$chunk_count.'.xml', 'w');

                $max_cell = chr(64+$header_length). ($chunk_size + 1);

                $xml_data = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><sheetPr><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr><dimension ref="A1:'.$max_cell.'"/><sheetViews><sheetView tabSelected="0" workbookViewId="0" showGridLines="true" showRowColHeaders="1"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="14.4" outlineLevelRow="0" outlineLevelCol="0"/><sheetData><row r="1" spans="1:'.$header_length.'">';

                for($i=0; $i<$header_length; $i++) {

                    $xml_data .= '<c r="'.chr(65+$i).'1" t="s"><v>'.$i.'</v></c>';

                }
        
                $xml_data .= '</row>';
                
                fwrite($xml, $xml_data);

                foreach($chunks as $i => $row) {

                    $row_content = $row_formatter($row);

                    $sharedStrings = '';

                    foreach($row_content as $content) {

                        if($content == 'SERIAL_NO' && $content != '0') {

                            $content = $i+1;
                        }

                        $sharedStrings .= '<si><t>'.$content.'</t></si>';
                    }
                    
                    fwrite($shared_string_file, $sharedStrings);

                    $xml_data = '<row r="'.($i+2).'" spans="1:'.$header_length.'">';

                    for($j=0; $j<$header_length; $j++) {

                        $xml_data .= '<c r="'.chr(65+$j).($i+2).'" t="s"><v>'.(++$shared_string_count).'</v></c>';

                    }

                    $xml_data .= '</row>';

                    fwrite($xml, $xml_data);

                }

                fwrite($xml, '</sheetData><printOptions gridLines="false" gridLinesSet="true"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="1" orientation="default" scale="100" fitToHeight="1" fitToWidth="1" pageOrder="downThenOver"/><headerFooter differentOddEven="false" differentFirst="false" scaleWithDoc="true" alignWithMargins="true"><oddHeader></oddHeader><oddFooter></oddFooter><evenHeader></evenHeader><evenFooter></evenFooter><firstHeader></firstHeader><firstFooter></firstFooter></headerFooter><tableParts count="0"/></worksheet>');

                fclose($xml);
            });

            fwrite($shared_string_file, '</sst>');

            fclose($shared_string_file);

            $zip = new ZipArchive();

            if ($zip->open($filename.".xlsx", ZipArchive::CREATE) === TRUE) {

                $zip->addFile($filename.'/[Content_Types].xml', '[Content_Types].xml');

                $zip->addFile($filename.'/docProps/app.xml', 'docProps/app.xml');

                $zip->addFile($filename.'/docProps/core.xml', 'docProps/core.xml');

                $zip->addFile($filename.'/xl/styles.xml', 'xl/styles.xml');

                $zip->addFile($filename.'/xl/theme/theme1.xml', 'xl/theme/theme1.xml');

                $zip->addFile($filename.'/xl/workbook.xml', 'xl/workbook.xml');

                $zip->addFile($filename.'/xl/sharedStrings.xml', 'xl/sharedStrings.xml');

                $zip->addFile($filename.'/xl/_rels/workbook.xml.rels', 'xl/_rels/workbook.xml.rels');

                $zip->addFile($filename.'/_rels/.rels', '_rels/.rels');

                $xml_files = glob($filename.'/xl/worksheets/*.xml');

                $rel_files = glob($filename.'/xl/worksheets/_rels/*.xml.rels');

                for($i=0; $i<$n_chunks; $i++) {

                    $zip->addFile($xml_files[$i], 'xl/worksheets/sheet'.($i+1).'.xml');

                    $zip->addFile($rel_files[$i], 'xl/worksheets/_rels/sheet'.($i+1).'.xml.rels');
                }

                $zip->close();

                static::deleteDir($filename);

                // echo("XLSX file created successfully.");
            } else {

                // echo("Failed to create XLSX file.");
            }
        } catch (Exception $e) {

            echo("Error in creating XLSX file: ".$e->getMessage());
        }

    }



    public static function deleteDir(string $dirPath): void {

        try{

            if (! is_dir($dirPath)) {
                
                throw new InvalidArgumentException("$dirPath must be a directory");
            }

            if (substr($dirPath, strlen($dirPath) - 1, 1) != '/') {

                $dirPath .= '/';
            }

            $files = glob($dirPath . '*', GLOB_MARK);

            foreach ($files as $file) {

                if (is_dir($file)) {

                    static::deleteDir($file);

                } else {

                    unlink($file);
                }
            }

            if(strpos($dirPath, '_rels') !== false)

                if(file_exists($dirPath.'.rels'))

                    unlink($dirPath.'.rels');

            rmdir($dirPath);

        } catch (Exception $e) {

            echo("Error in deleting directory: ".$e->getMessage());
        }
    }
    
}