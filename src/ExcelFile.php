<?php

namespace Sulabh\ExcelFilePackage;

use ZipArchive;
use Exception;
use InvalidArgumentException;

class ExcelFile
{
    private static $lowercase_headers = null;
    private static $excel_column_index = [];
    private static $column_index = [];
    public static function createExcelFile($data, $chunk_size, $headers, $filename, $row_formatter, $total_data_count = 0)
    {
        try {
            $total_data_count = $total_data_count ?? (is_countable($data) ? $data->count() : count($data));

            $basePath = dirname(__DIR__, 4); 
            $publicPath = $basePath . '/public';
            
            if (!file_exists($publicPath.'/exports')) {
                mkdir($publicPath.'/exports', 0777, true);
            }

            $filename = str_replace('.xlsx', '', $filename);
            $filename = $publicPath.'/exports/'.basename($filename);

            // Create directory structure
            $dirs = [
                $filename,
                $filename.'/docProps',
                $filename.'/_rels',
                $filename.'/xl',
                $filename.'/xl/_rels',
                $filename.'/xl/theme',
                $filename.'/xl/worksheets',
                $filename.'/xl/worksheets/_rels'
            ];

            foreach ($dirs as $dir) {
                if (!file_exists($dir)) {
                    mkdir($dir, 0777, true);
                }
            }

            $chunk_count = 0;
            $n_chunks = ceil($total_data_count / $chunk_size);

            // Generate [Content_Types].xml
            $content_type_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';

            for($i = 1; $i <= $n_chunks; $i++) {
                $content_type_xml .= '<Override PartName="/xl/worksheets/sheet'.$i.'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
            }
            
            $content_type_xml .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>';

            file_put_contents($filename.'/[Content_Types].xml', $content_type_xml);

            // Generate docProps/app.xml
            $docProcs_app_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>Microsoft Excel</Application>
    <DocSecurity>0</DocSecurity>
    <ScaleCrop>false</ScaleCrop>
    <HeadingPairs>
        <vt:vector size="2" baseType="variant">
            <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
            <vt:variant><vt:i4>'.$n_chunks.'</vt:i4></vt:variant>
        </vt:vector>
    </HeadingPairs>
    <TitlesOfParts>
        <vt:vector size="'.$n_chunks.'" baseType="lpstr">';

            for($i = 1; $i <= $n_chunks; $i++) {
                $docProcs_app_xml .= '<vt:lpstr>Sheet'.$i.'</vt:lpstr>';
            }

            $docProcs_app_xml .= '</vt:vector>
    </TitlesOfParts>
    <Company></Company>
    <LinksUpToDate>false</LinksUpToDate>
    <SharedDoc>false</SharedDoc>
    <HyperlinksChanged>false</HyperlinksChanged>
    <AppVersion>16.0300</AppVersion>
</Properties>';

            file_put_contents($filename.'/docProps/app.xml', $docProcs_app_xml);

            // Generate xl/workbook.xml
            $xl_workbook_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
    <workbookPr defaultThemeVersion="124226"/>
    <bookViews>
        <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>
    </bookViews>
    <sheets>';

            for($i = 1; $i <= $n_chunks; $i++) {
                $xl_workbook_xml .= '<sheet name="Sheet'.$i.'" sheetId="'.$i.'" r:id="rId'.($i+3).'"/>';
            }

            $xl_workbook_xml .= '</sheets>
    <calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>';

            file_put_contents($filename.'/xl/workbook.xml', $xl_workbook_xml);

            // Generate xl/_rels/workbook.xml.rels
            $xl_rels_workbook_xml_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';

            for($i = 1; $i <= $n_chunks; $i++) {
                $xl_rels_workbook_xml_rels .= '<Relationship Id="rId'.($i+3).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'.$i.'.xml"/>';
            }

            $xl_rels_workbook_xml_rels .= '</Relationships>';

            file_put_contents($filename.'/xl/_rels/workbook.xml.rels', $xl_rels_workbook_xml_rels);

            // Generate _rels/.rels
            file_put_contents($filename.'/_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>');

            // Generate docProps/core.xml
            $docProcs_core_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>Excel Generator</dc:creator>
    <cp:lastModifiedBy>Excel Generator</cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">'.date('Y-m-d\TH:i:s\Z').'</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">'.date('Y-m-d\TH:i:s\Z').'</dcterms:modified>
</cp:coreProperties>';

            file_put_contents($filename.'/docProps/core.xml', $docProcs_core_xml);

            // Generate xl/theme/theme1.xml
            $xl_theme_theme1_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <a:themeElements>
        <a:clrScheme name="Office">
            <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
            <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
            <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
            <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
            <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
            <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
            <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
            <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
            <a:accent6><a:srgbClr val="F79646"/></a:accent6>
            <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
            <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Office">
            <a:majorFont>
                <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
                <a:font script="Jpan" typeface="ＭＳ ゴシック"/>
                <a:font script="Hang" typeface="맑은 고딕"/>
                <a:font script="Hans" typeface="等线"/>
                <a:font script="Hant" typeface="新細明體"/>
                <a:font script="Arab" typeface="Times New Roman"/>
                <a:font script="Hebr" typeface="Times New Roman"/>
                <a:font script="Thai" typeface="Angsana New"/>
                <a:font script="Ethi" typeface="Nyala"/>
                <a:font script="Beng" typeface="Vrinda"/>
                <a:font script="Gujr" typeface="Shruti"/>
                <a:font script="Khmr" typeface="MoolBoran"/>
                <a:font script="Knda" typeface="Tunga"/>
                <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                <a:font script="Thaa" typeface="MV Boli"/>
                <a:font script="Deva" typeface="Mangal"/>
                <a:font script="Telu" typeface="Gautami"/>
                <a:font script="Taml" typeface="Latha"/>
                <a:font script="Orya" typeface="Kalinga"/>
                <a:font script="Mlym" typeface="Kartika"/>
                <a:font script="Laoo" typeface="DokChampa"/>
                <a:font script="Sinh" typeface="Iskoola Pota"/>
                <a:font script="Mong" typeface="Mongolian Baiti"/>
                <a:font script="Viet" typeface="Times New Roman"/>
                <a:font script="Uigh" typeface="Microsoft Uighur"/>
                <a:font script="Geor" typeface="Sylfaen"/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="Calibri" panose="020F0502020204030204"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
                <a:font script="Jpan" typeface="ＭＳ 明朝"/>
                <a:font script="Hang" typeface="맑은 고딕"/>
                <a:font script="Hans" typeface="等线"/>
                <a:font script="Hant" typeface="新細明體"/>
                <a:font script="Arab" typeface="Arial"/>
                <a:font script="Hebr" typeface="Arial"/>
                <a:font script="Thai" typeface="Angsana New"/>
                <a:font script="Ethi" typeface="Nyala"/>
                <a:font script="Beng" typeface="Vrinda"/>
                <a:font script="Gujr" typeface="Shruti"/>
                <a:font script="Khmr" typeface="DaunPenh"/>
                <a:font script="Knda" typeface="Tunga"/>
                <a:font script="Syrc" typeface="Estrangelo Edessa"/>
                <a:font script="Cher" typeface="Plantagenet Cherokee"/>
                <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
                <a:font script="Tibt" typeface="Microsoft Himalaya"/>
                <a:font script="Thaa" typeface="MV Boli"/>
                <a:font script="Deva" typeface="Mangal"/>
                <a:font script="Telu" typeface="Gautami"/>
                <a:font script="Taml" typeface="Latha"/>
                <a:font script="Orya" typeface="Kalinga"/>
                <a:font script="Mlym" typeface="Kartika"/>
                <a:font script="Laoo" typeface="DokChampa"/>
                <a:font script="Sinh" typeface="Iskoola Pota"/>
                <a:font script="Mong" typeface="Mongolian Baiti"/>
                <a:font script="Viet" typeface="Arial"/>
                <a:font script="Uigh" typeface="Microsoft Uighur"/>
                <a:font script="Geor" typeface="Sylfaen"/>
            </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                        <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
                        <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="1"/></a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
                        <a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
                        <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="0"/></a:gradFill>
            </a:fillStyleLst>
            <a:lnStyleLst>
                <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill>
                    <a:prstDash val="solid"/></a:ln>
                <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                    <a:prstDash val="solid"/></a:ln>
                <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                    <a:prstDash val="solid"/></a:ln>
            </a:lnStyleLst>
            <a:effectStyleLst>
                <a:effectStyle><a:effectLst/></a:effectStyle>
                <a:effectStyle><a:effectLst/></a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
            </a:effectStyleLst>
            <a:bgFillStyleLst>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
                        <a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
                        <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
                    </a:gsLst>
                    <a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>
                </a:gradFill>
            </a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
    <a:objectDefaults/>
    <a:extraClrSchemeLst/>
</a:theme>';

            file_put_contents($filename.'/xl/theme/theme1.xml', $xl_theme_theme1_xml);

            // Generate xl/styles.xml
            $xl_styles_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="0"/>
    <fonts count="1">
        <font>
            <sz val="11"/>
            <color rgb="FF000000"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="2">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
    </fills>
    <borders count="1">
        <border>
            <left/><right/><top/><bottom/>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>';

            file_put_contents($filename.'/xl/styles.xml', $xl_styles_xml);

            $sharedStrings = sprintf(
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">',
                $total_data_count * count($headers) + count($headers),
                count($headers)
            );
            
            foreach ($headers as $header) {
                $sharedStrings .= sprintf('<si><t>%s</t></si>', htmlspecialchars($header, ENT_XML1));
            }

            touch($filename.'/xl/sharedStrings.xml');

            $shared_string_file = fopen($filename.'/xl/sharedStrings.xml', 'a');

            fwrite($shared_string_file, $sharedStrings);
            
            $shared_string_count = count($headers);

            foreach ($headers as $i=>$header) {
                static::$column_index[$header] = $i;
            }

            // Process data in chunks
            $data->chunk($chunk_size, function($chunks) use ($filename, &$chunk_count, &$shared_string_file, $headers, $row_formatter, &$shared_string_count) {
                $chunk_count++;

                $worksheet_file = $filename.'/xl/worksheets/sheet'.$chunk_count.'.xml';

                touch($worksheet_file);

                $xml = fopen($worksheet_file, 'a');

                $header_length = count($headers);
                $max_cell = static::numberToExcelColumn($header_length) . strval(count($chunks) + 1);

                // Create worksheet XML
                $worksheet_xml = sprintf(
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                        <dimension ref="A1:%s"/>
                        <sheetViews>
                            <sheetView workbookViewId="0">
                                <selection activeCell="A1" sqref="A1"/>
                            </sheetView>
                        </sheetViews>
                        <sheetFormatPr defaultRowHeight="15"/>
                        <sheetData>
                            <row r="1" spans="1:%d">',
                    $max_cell,
                    $header_length
                );
                // Add headers
                $worksheet_headers = [];
                for($i = 0; $i < $header_length; $i++) {
                    $worksheet_headers[$i] = sprintf('<c r="%s1" t="s"><v>%d</v></c>', static::numberToExcelColumn(1 + $i), $i);
                }
                
                $worksheet_xml = $worksheet_xml . implode('', $worksheet_headers). '</row>';

                fwrite($xml, $worksheet_xml);

                // Add data rows
                foreach($chunks as $index => $row) {
                    $row_data = $row_formatter($row);
                    
                    $sharedStrings = [];

                    $worksheet_xml = [];
                    $i = 0;
                    $j = 0;

                    $worksheet_xml[$i++] = sprintf('<row r="%d" spans="1:%d">', $index + 2, $header_length);
                    
                    foreach($row_data as $col => $value) {
                        $col_index = is_numeric($col) ? $col : static::$column_index[$col];
                        if($value === 'SERIAL_NO') {
                            $value = $index + 1;
                        }

                        $sharedStrings[$j++] = "<si><t>$value</t></si>";
                        $worksheet_xml[$i++] = sprintf('<c r="%s%d" t="s"><v>%d</v></c>', static::numberToExcelColumn(1 + $col_index), $index + 2, $shared_string_count++);
                    }
                    fwrite($shared_string_file, implode('', $sharedStrings));
                    $worksheet_xml[$i++] = '</row>';

                    fwrite($xml, implode('', $worksheet_xml));
                }

                $worksheet_xml = '</sheetData>
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>';

                fwrite($xml, $worksheet_xml);
                fclose($xml);

                // Create empty worksheet rels
                file_put_contents($filename.'/xl/worksheets/_rels/sheet'.$chunk_count.'.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
            });

            fwrite($shared_string_file, '</sst>');
            fclose($shared_string_file);

            // Create ZIP archive
            $zip = new ZipArchive();
            $zipFilename = $filename.'.xlsx';

            if ($zip->open($zipFilename, ZipArchive::CREATE | ZipArchive::OVERWRITE) === TRUE) {
                // Add all required files to the zip
                $files = [
                    '[Content_Types].xml',
                    '_rels/.rels',
                    'docProps/app.xml',
                    'docProps/core.xml',
                    'xl/workbook.xml',
                    'xl/_rels/workbook.xml.rels',
                    'xl/styles.xml',
                    'xl/theme/theme1.xml',
                    'xl/sharedStrings.xml'
                ];

                foreach ($files as $file) {
                    $zip->addFile($filename.'/'.$file, $file);
                }

                // Add worksheets
                for ($i = 1; $i <= $n_chunks; $i++) {
                    $zip->addFile($filename.'/xl/worksheets/sheet'.$i.'.xml', 'xl/worksheets/sheet'.$i.'.xml');
                    $zip->addFile($filename.'/xl/worksheets/_rels/sheet'.$i.'.xml.rels', 'xl/worksheets/_rels/sheet'.$i.'.xml.rels');
                }

                $zip->close();

                // Clean up temporary files
                static::deleteDir($filename);

                return true;
            } else {
                throw new Exception("Failed to create ZIP archive");
            }
        } catch (Exception $e) {
            // Clean up in case of error
            if (isset($filename) && file_exists($filename)) {
                static::deleteDir($filename);
            }
            throw $e;
        }
    }

    private static function columnToIndex(string $column_name, array $headers): int
    {
        if (self::$lowercase_headers === null) {
            self::$lowercase_headers = array_map('strtolower', $headers);
        }
    
        $column_name = strtolower($column_name);
        $index = array_search($column_name, self::$lowercase_headers);
    
        if ($index === false) {
            throw new InvalidArgumentException("Column $column_name not found in headers");
        }
    
        return $index;
    }
    

    private static function numberToExcelColumn(int $number): string
    {
        $column = '';
        while ($number > 0) {
            $mod = ($number - 1) % 26;
            if(!isset(self::$excel_column_index[$mod])) {
                self::$excel_column_index[$mod] = chr(65 + $mod);
            }
            $column = self::$excel_column_index[$mod] . $column;
            $number = intdiv($number, 26);
        }
        return $column;
    }

    public static function deleteDir(string $dirPath): void
    {
        try {
            // Check if directory exists
            if (!is_dir($dirPath)) {
                throw new InvalidArgumentException("$dirPath must be a directory");
            }
    
            // Normalize the directory path (remove trailing slashes if any)
            $dirPath = rtrim($dirPath, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR;
    
            // Use RecursiveIteratorIterator for efficient directory traversal
            $iterator = new \RecursiveIteratorIterator(
                new \RecursiveDirectoryIterator($dirPath, \RecursiveDirectoryIterator::SKIP_DOTS),
                \RecursiveIteratorIterator::CHILD_FIRST
            );
    
            // Delete all files and directories
            foreach ($iterator as $fileinfo) {
                if ($fileinfo->isDir()) {
                    rmdir($fileinfo->getRealPath());
                } else {
                    unlink($fileinfo->getRealPath());
                }
            }
    
            // If the directory contains _rels, remove the corresponding file
            if (strpos($dirPath, '_rels') !== false && file_exists($dirPath . '.rels')) {
                unlink($dirPath . '.rels');
            }
    
            // Finally, remove the main directory
            rmdir($dirPath);
        } catch (Exception $e) {
            // Log the error instead of printing
            error_log("Error in deleting directory: " . $e->getMessage());
        }
    }    
}