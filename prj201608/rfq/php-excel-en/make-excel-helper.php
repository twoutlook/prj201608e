<?php

/** Error reporting */
error_reporting(E_ALL);

$tool = new MarkTool();
echo '&lt;?php <br>';
$tool->makePercentFormat();
$tool->makeSum(); //
//
//
//
//$tool->makeRmbStyle(); //makeRmbStyle
echo "<br> // RMB<br> ";
//$moneyArrRMB = '[{"items":[19, 20, 21, 22, 23,24, 31, 33, 36, 37, 38, 39, 42, 45, 46, 47, 49, 52, 53, 56, 57, 60, 63, 64, 65, 68, 70, 72, 73, 76, 77, 81, 83, 87, 88, 91, 94, 95, 98, 99, 102, 103, 104, 105, 107, 108, 109, 110, 111]}]';
// to 70

$moneyArrRMB = '[{"items":[19, 20, 21, 22, 23,24, 31, 33, 36, 37, 38, 39, 42, 45, 46, 47, 49, 52, 53, 56, 57, 60, 63, 64, 65, 68, 70,  73,74,77,78,82,84,88,89,90,91,92,95,98,99,102,103,106,107,108,109,111,112,113,114,115,117,118]}]';
$tool->makeMoneyStyle("¥", $moneyArrRMB);
echo "<br> // USD<br> ";
$moneyArrUSD = '[{"items":[25,116]}]';
$tool->makeMoneyStyle("$", $moneyArrUSD);

echo "<br> // USD 公式計算<br> ";
$tool->extendColToCDEFGH(25,"=C24/6.35");
$tool->extendColToCDEFGH(116,"=C115/6.35");

//32,=C11
$tool->extendColToCDEFGH(32,"=C11");
//33,=C31*C32/1000
$tool->extendColToCDEFGH(33,"=C31*C32/1000");

//35,=100*C32*C16/(C32*C16+C34)
//使用 Excel 百分比 --- 35,=C32*C16/(C32*C16+C34)
$tool->extendColToCDEFGH(35,"=C32*C16/(C32*C16+C34)");

//37,                        =(C31-C36)*C34/1000/C16
$tool->extendColToCDEFGH(37,"=(C31-C36)*C34/1000/C16");

//38,                        =(C32+C34)*C31*0.02/1000/C16
$tool->extendColToCDEFGH(38,"=(C32+C34)*C31*0.02/1000/C16");

//39,                        =IF(ISNA(C33+C37+C38),0,(C33+C37+C38))
$tool->extendColToCDEFGH(39,"=IF(ISNA(C33+C37+C38),0,(C33+C37+C38))");

//44,                        =3600/C43
$tool->extendColToCDEFGH(44,"=3600/C43");

//45,                       =C42/C44
$tool->extendColToCDEFGH(45,"=C42/C44");

//49,                       =IF(ISNA((C45+C46)*(1+(1-C48/100))/C16),0,((C45+C46)*(1+(1-C48/100))/C16))
// fix percent              =IF(ISNA((C45+C46)*(1+(1-C48))/C16),0,((C45+C46)*(1+(1-C48))/C16))
$tool->extendColToCDEFGH(49,"=IF(ISNA((C45+C46)*(1+(1-C48))/C16),0,((C45+C46)*(1+(1-C48))/C16))");




//53,                        =(C51/3600)*C52
$tool->extendColToCDEFGH(53,"=(C51/3600)*C52");

//57,                        =(C55/3600)*C56
$tool->extendColToCDEFGH(57,"=(C55/3600)*C56");

//60,                        =C57*(1+(1-C58/100))
// fix percent               =C57*(1+(1-C58))
$tool->extendColToCDEFGH(60,"=C57*(1+(1-C58))");

//64,  65                    =(C62/3600)*C63
$tool->extendColToCDEFGH(64,"=(C62/3600)*C63");
$tool->extendColToCDEFGH(65,"=(C62/3600)*C63");


// 70, =IF(ISNA((C67/3600)*C68 * (1 + (1 - C69 / 100))),0,((C67/3600)*C68 * (1 + (1 - C69 / 100))))
// fix percent               =IF(ISNA((C67/3600)*C68 * (1 + (1 - C69 ))),0,((C67/3600)*C68 * (1 + (1 - C69 ))))
$tool->extendColToCDEFGH(70,"=IF(ISNA((C67/3600)*C68 * (1 + (1 - C69 ))),0,((C67/3600)*C68 * (1 + (1 - C69 ))))");

// 74                        =IF(ISNA((C72/3600)*C73),0,((C72/3600)*C73))
$tool->extendColToCDEFGH(74,"=IF(ISNA((C72/3600)*C73),0,((C72/3600)*C73))");

//78                         =IF(ISNA(C77),0,(C76/3600)*C77)
$tool->extendColToCDEFGH(78,"=IF(ISNA(C77),0,(C76/3600)*C77)");

// 84                        =C81*C82*C83
$tool->extendColToCDEFGH(84,"=C81*C82*C83");

// 95                        =C87*(SUM(C88:C92))*(1+(1-C93/100))*C94
// fix percent               =C87*(SUM(C88:C92))*(1+(1-C93/100))*C94
//$tool->extendColToCDEFGH(95,"=C87*(SUM(C88:C92))*(1+(1-C93))*C94");

echo "//[[A0607]] 加入的 遮蔽費用等三項, 公式之前有誤";
$tool->extendColToCDEFGH(95,"=(C87*(SUM(C88:C89))+SUM(C90:C92))*(1+(1-C93))*C94");





// 99                        =C98
$tool->extendColToCDEFGH(99,"=C98");

// 103                     =C101/3600*C102
$tool->extendColToCDEFGH(103,"=C101/3600*C102");

// 108                        =C105/3600*C106+C107
//$tool->extendColToCDEFGH(108,"=C105/3600*C106+C107");

//echo "// [[A0607]] FIX 106, 108 FORMULA IN EXCEL";
echo "//[[A0613]] #106 BUG";
// 108                        =C105/3600*C106+C107


//[[A0613]]
//$tool->extendColToCDEFGH(106,"=C105/3600");//BUG
$tool->extendColToCDEFGH(106,"=C105*35/3600");

$tool->extendColToCDEFGH(108,"=C106+C107");




// 111                        =C109*C110/100
// fix percent                =C109*C110
$tool->extendColToCDEFGH(111,"=C109*C110");

// 117                        =C115*1.17
$tool->extendColToCDEFGH(117,"=C115*1.17");


/*
  $tool->makeCell32(32);
  $tool->extendCell34X(34, "=100*C31*C16/(C31*C16+C33)/100"); //注意 EXCEL 和 ENTERPRISESHEET 的百分比表達方式不同
  $tool->extendCell34X(36, "=(C30-C35)*C33/1000/C16");
  $tool->extendCell34X(37, "=(C31+C33)*C30*0.02/1000/C16");
  $tool->extendCell34X(38, "=IF(ISNA(C32+C36+C37),0,(C32+C36+C37))");
  //
  $tool->extendCell34X(43, "=3600/C42");
  $tool->extendCell34X(44, "=C41/C43 ");
  //$tool->extendCell34X(45, "  ");
  //$tool->extendCell34X(46, "  ");
  // 47 百分比要處理
  // $tool->extendCell34X(48, "=(C44+C45)*(1+(1-C47/100))/C16");  // before fix
  $tool->extendCell34X(48, "=(C44+C45)*(1+(1-C47))/C16");  //  after fix

  $tool->extendCell34X(52, "=(C50/3600)*C51");
  //
  $tool->extendCell34X(56, "=(C54/3600)*C55");
  //=C56*(1+(1-C57/100))
  //$tool->extendCell34X(59, "=C56*(1+(1-C57/100))"); // TO FIX PERCENT
  $tool->extendCell34X(59, "=C56*(1+(1-C57))");

  // 63,64 SAME =(C61/3600)*C62'
  $tool->extendCell34X(63, "=(C61/3600)*C62");
  $tool->extendCell34X(64, "=(C61/3600)*C62");
  //
  //$tool->extendCell34X(69, "=(C66/3600)*C67 * (1 + (1 - C68 / 100))");
  $tool->extendCell34X(69, "=IF(ISNA((C66/3600)*C67 * (1 + (1 - C68))),0,(C66/3600)*C67 * (1 + (1 - C68)))");

  //73
  //(C71/3600)*C72
  $tool->extendCell34X(73, "=IF(ISNA((C71/3600)*C72),0,(C71/3600)*C72)");

  //77
  //"=IF(ISNA(C76),0,(C75/3600)*C76)
  $tool->extendCell34X(77, "=IF(ISNA(C76),0,(C75/3600)*C76)");

  //83
  //
  $tool->extendCell34X(83, "=C80*C81*C82");
  //91
  //"=C86*(C87+C88)*(1+(1-C89/100))*C90
  //$tool->extendCell34X(91, "=C86*(C87+C88)*(1+(1-C89/100))*C90");
  $tool->extendCell34X(91, "=C86*(C87+C88)*(1+(1-C89))*C90");

  //95
  //
  $tool->extendCell34X(95, "=C94");

  //99
  //
  $tool->extendCell34X(99, "=C98");

  //104
  //
  $tool->extendCell34X(104, "=C102+C103");

  //107
  //
  //$tool->extendCell34X(107, "=C105*C106/100");
  $tool->extendCell34X(107, "=C105*C106");


  //110
  //
  $tool->extendCell34X(110, "=C108+C109");

 */


//
//
//
//
//
//    var colorStep = "#A9BCF5";
//    var colorStepEnd = "#E6E6E6";
//    var colorSect = "#837E7C"; //bgc: colorSect, fm: "money|¥|2|none", dsd: "ed", cal: true
//    var colorDdl = "#F9E79F"; //#82E0AA  
//    var colorInput = "#F4D03F"; // 
//    var arrStepEnd = [23, 24, 38, 48, 52, 59, 64, 69, 73, 77, 83, 91, 95, 99, 104, 105, 110, 111, 112];
$fontJsonStrItalicTrue = '{"0000A0":[31,36,42,46,47,52,56,63,68,73,77,82,102,106]}'; //直接查表或是立即計算的
$colorJsonStrStepStart = '{"A9BCF5":[15,28,29,40,50,54,61,66,71,75,79,85,96,100,104,110,112,7]}';

$colorJsonStrStepEnd = '{"BE6E6E6":[39,49,53,60,65,70,74,78,84,95,99,103,108,111,114]}';
$jsonDdl = '{"F9E79F":[10,12,30,41,66,71,75,80,86]}';

$jsonInput = '{"F4D03F":[11,16,19,20,21,22,23,34,43,48,51,55,58,62,67,69,72,76,81,83,87,88,89,90,91,92,93,94,97,98,101,105,107,110,112,113]}';
//$jsonInput = '{"F4D03F":[11,16]}';

$jsonBigTotal = '{"cccccc":[24,25,109,115,116,117]}';

//[[A0601]]
//    var colorVersion = "#98AFC7"; // 
$jsonVersion = '{"98AFC7":[7,8,9]}';

//$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
//>getFont()->setItalic(true);  

$tool->makeColorFillStyle("A", $colorJsonStrStepStart);
$tool->makeColorFillStyle("B", $colorJsonStrStepStart);
$tool->makeColorFillStyle("B", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("C", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("D", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("E", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("F", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("G", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("H", $colorJsonStrStepEnd);
$tool->makeColorFillStyle("C", $jsonDdl);
$tool->makeColorFillStyle("D", $jsonDdl);
$tool->makeColorFillStyle("E", $jsonDdl);
$tool->makeColorFillStyle("F", $jsonDdl);
$tool->makeColorFillStyle("G", $jsonDdl);
$tool->makeColorFillStyle("H", $jsonDdl);
$tool->makeColorFillStyle("C", $jsonInput);
$tool->makeColorFillStyle("D", $jsonInput);
$tool->makeColorFillStyle("E", $jsonInput);
$tool->makeColorFillStyle("F", $jsonInput);
$tool->makeColorFillStyle("G", $jsonInput);
$tool->makeColorFillStyle("H", $jsonInput);

$tool->makeColorFillStyle("A", $jsonBigTotal);
$tool->makeColorFillStyle("B", $jsonBigTotal);
$tool->makeColorFillStyle("C", $jsonBigTotal);
$tool->makeColorFillStyle("D", $jsonBigTotal);
$tool->makeColorFillStyle("E", $jsonBigTotal);
$tool->makeColorFillStyle("F", $jsonBigTotal);
$tool->makeColorFillStyle("G", $jsonBigTotal);
$tool->makeColorFillStyle("H", $jsonBigTotal);

echo "<br><br>//[[A0601]] 顯示版本信息";
$tool->makeColorFillStyle("A", $jsonVersion);

echo "<br><br>//[[A0601]] 加方格線";
$tool->makeCellsBorderColRowFromTo("ABCEDFGH", 3, 119);


$tool->makeFontItalic("C",$fontJsonStrItalicTrue);
$tool->makeFontItalic("D",$fontJsonStrItalicTrue);
$tool->makeFontItalic("E",$fontJsonStrItalicTrue);
$tool->makeFontItalic("F",$fontJsonStrItalicTrue);
$tool->makeFontItalic("G",$fontJsonStrItalicTrue);
$tool->makeFontItalic("H",$fontJsonStrItalicTrue);


class MarkTool {
    /*
      $objPHPExcel->getActiveSheet()
      ->setCellValue('C23', '=SUM(C19:C22)')
      ->setCellValue('D23', '=SUM(D19:D22)')
      ->setCellValue('E23', '=SUM(E19:E22)')
      ;
     */

    public function makeCell32($row) {
        echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";

        $colNameArr = Array("C", "D", "E", "F", "G", "H");
        $str = " <br> \$objPHPExcel->getActiveSheet() <br>";
        for ($i = 0; $i < 6; $i++) {
            $colName = $colNameArr[$i];
            //=C30*C31/1000
            $colRow = $colName . $row;
            $cell = "=" . $colName . "30*" . $colName . "31/1000";
            $str.="  ->setCellValue('$colRow', '$cell') <br>";
        }
        echo $str . ";";
    }

    //=100*C31*C16/(C31*C16+C33)
//      public function makeCell34X($rowNum,$cellFormula) {
//        $colNameArr = Array("C", "D", "E", "F", "G", "H");
//        $str = " <br> \$objPHPExcel->getActiveSheet() <br>";
//        for ($i = 0; $i < 6; $i++) {
//            $colName = $colNameArr[$i];
//            //=C30*C31/1000
//            $colRow = $colName . $row;
//            $cell = "=" . $colName . "30*" . $colName . "31/1000";
//            $str.="  ->setCellValue('$colRow', '$cell') <br>";
//        }
//        echo $str . ";";
//    }


    public function extendColToCDEFGH($rowNum, $cellFormula) {
        echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";

        echo "<br><br>// ---  extendColToCDEFGH($rowNum,$cellFormula) ---<br>";

        $seq = "0ABCDEFGH";
        echo " <br> \$objPHPExcel->getActiveSheet() <br>";
        echo "->setCellValue('C$rowNum', '$cellFormula')<br>";
        for ($k = 4; $k <= 8; $k++) {
            //    $strD = str_replace("col: 3", "col: $k", $cellFormula);
            $colName = substr($seq, $k, 1);
            $newCellformula = $cellFormula;
            for ($item = 1; $item < 115; $item++) {
                $oldStr = "C$item";
                $newStr = $colName . $item;
                $newCellformula = str_replace($oldStr, $newStr, $newCellformula);
            }
            echo "->setCellValue('$colName$rowNum', '$newCellformula')<br>";
        }
        echo ";<br>";
    }

    
    
    public function extendCell34X($rowNum, $cellFormula) {
        echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";

        echo "<br><br>// ---  extendCell34X($rowNum,$cellFormula) ---<br>";

        $seq = "0ABCDEFGH";
        echo " <br> \$objPHPExcel->getActiveSheet() <br>";
        echo "->setCellValue('C$rowNum', '$cellFormula')<br>";
        for ($k = 4; $k <= 8; $k++) {
            //    $strD = str_replace("col: 3", "col: $k", $cellFormula);
            $colName = substr($seq, $k, 1);
            $newCellformula = $cellFormula;
            for ($item = 1; $item < 115; $item++) {
                $oldStr = "C$item";
                $newStr = $colName . $item;
                $newCellformula = str_replace($oldStr, $newStr, $newCellformula);
            }
            echo "->setCellValue('$colName$rowNum', '$newCellformula')<br>";
        }
        echo ";<br>";
    }

    /*
      public function makeUsd() {
      echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";

      $strUsd = '[{"usd":24, "rmb":23},{"usd":112, "rmb":111}]';
      $objUsd = json_decode($strUsd);
      //        print_r($objUsd);

      $str = " <br> \$objPHPExcel->getActiveSheet() <br>";
      foreach ($objUsd as $key => $obj) {
      $colNameArr = Array("C", "D", "E", "F", "G", "H");
      for ($i = 0; $i < 6; $i++) {
      $colName = $colNameArr[$i];
      $str.="  ->setCellValue('" . $colName . $obj->usd . "', '=" . $colName . $obj->rmb . "/6.35') <br>";
      }
      }
      echo $str . ";";
      }
     */

    public function makeSum() {
        echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";


        $strUsd = '[{"sum":24, "items":[19,20,21,22,23]},{"sum":115, "items":[109,111,114]},{"sum":114, "items":[112,113]},{"sum":109, "items":[39,49,53,60,65,70,74,78,84,95,99,103,108]}]';
        $objUsd = json_decode($strUsd);
//        print_r($objUsd);
//__LINE__、__FUNCTION__


        $str = " <br> \$objPHPExcel->getActiveSheet() <br>";

        foreach ($objUsd as $key => $obj) {
            $colNameArr = Array("C", "D", "E", "F", "G", "H");
            for ($i = 0; $i < 6; $i++) {
                $colName = $colNameArr[$i];
//                $str.="  ->setCellValue('" . $colName . $obj->sum . "', '=" . $colName . $obj->rmb . "/6.35') <br>";
                $str.="  ->setCellValue('" . $colName . $obj->sum . "', '=";

                $items = $obj->items;
                for ($j = 0; $j < count($items); $j++) {
                    $str.= $colName . $items[$j];
                    if ($j < count($items) - 1) {
                        $str.="+";
                    }
                }
                $str.="')<br>";
            }
        }
        echo $str . ";<br><br>";
    }



    public function makeRmbStyle() {
        echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";

        $strRmb = '[{"items":[19,20,21,22,23,32, 38, 48, 52, 59, 64, 69, 73, 77, 83, 91, 95, 99, 104, 105, 110, 111]}]';
        $objRmb = json_decode($strRmb);
//        print_r($objRmb);



        foreach ($objRmb as $key => $obj) {

            $items = $obj->items;
            for ($j = 0; $j < count($items); $j++) {
//                $str = " <br> \$objPHPExcel->getActiveSheet()->duplicateConditionalStyle( <br>";
//                $str.= "  \$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C" . $items[$j] . ":H" . $items[$j];
                $str = "\$objPHPExcel->getActiveSheet()->getStyle('C" . $items[$j] . ":H" . $items[$j] . "')->getNumberFormat()->setFormatCode(\"¥#,##0.00\")";
                echo $str . ";<br>";
            }
        }
    }

    public function makePercentFormat() {
        echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";

        echo "<br>// makePercentFormat 0.00% ok, FORMAT_PERCENTAGE_00 為什麼不行?<br>";
        $strRmb = '[{"items":[35,48,58,69,93,110]}]';
        $objRmb = json_decode($strRmb);

        foreach ($objRmb as $key => $obj) {

            $items = $obj->items;
            for ($j = 0; $j < count($items); $j++) {
//                $str = " <br> \$objPHPExcel->getActiveSheet()->duplicateConditionalStyle( <br>";
//                $str.= "  \$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C" . $items[$j] . ":H" . $items[$j];
                $str = "\$objPHPExcel->getActiveSheet()->getStyle('C" . $items[$j] . ":H" . $items[$j] . "')->getNumberFormat()->setFormatCode(\"0.00%\")";
                echo $str . ";<br>";
            }
        }
    }

    public function makeMoneyStyle($money, $moneyArr) {
        echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";

        $objMoney = json_decode($moneyArr);
//        print_r($objRmb);



        foreach ($objMoney as $key => $obj) {

            $items = $obj->items;
            for ($j = 0; $j < count($items); $j++) {
//                $str = " <br> \$objPHPExcel->getActiveSheet()->duplicateConditionalStyle( <br>";
//                $str.= "  \$objPHPExcel->getActiveSheet()->getStyle('C19')->getConditionalStyles(), 'C" . $items[$j] . ":H" . $items[$j];
                $str = "\$objPHPExcel->getActiveSheet()->getStyle('C" . $items[$j] . ":H" . $items[$j] . "')->getNumberFormat()->setFormatCode(\"" . $money . "#,##0.00\")";
                echo $str . ";<br>";
            }
        }
    }

    //[[A0601]]
    public function makeCellsBorderColRowFromTo($colName, $rowFrom, $rowTo) {
        echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";

        $str = "<br><br>\$styleArr = array( ";
        $str .="'borders' => array(";
        $str .="    'allborders' => array(";
        $str .="        'style' => 'thin'"; //, //细边框  //   const BORDER_THIN             = 'thin';
        $str .="    )";
        $str .=")";
        $str .=");<br>";
        for ($k = 0; $k < strlen($colName); $k++) {
            for ($i = $rowFrom; $i <= $rowTo; $i++) {
                $cell = substr($colName, $k, 1) . $i;
                $str .= "\$objPHPExcel->getActiveSheet()->getStyle('$cell')->applyFromArray(\$styleArr);<br>"; //这里就是画出从单元格A5到N5的边框，看单元格最右边在哪哪个格就把这个N改
            }
        }
        echo $str;
    }

    // cell format
    // http://www.cnblogs.com/freespider/p/3284828.html
    // for column B only
    public function makeColorFillStyle($col, $str) {
        echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";

        echo "<br><br>// ---  makeColorFillStyle($col, $str)---<br> ";
        $json = json_decode($str);
//        print_r($json);
        //$objPHPExcel->getActiveSheet()->getStyle('A1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
//        ->getStartColor()->setARGB('FFA9BCF5');

        foreach ($json as $key => $val) {
//            echo $key;
//            print_r($val);
            foreach ($val as $item) {
//                 echo $item;
                $str = "\$objPHPExcel->getActiveSheet()->getStyle('$col$item')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)";
                $str.=" ->getStartColor()->setARGB('$key');<br>";
                echo $str;
            }
        }
    }

    
     public function makeFontItalic($col, $str) {
        echo "<br><br>//<br>// file:" . __FILE__ . " line:" . __LINE__ . " function: " . __FUNCTION__ . "<br>//<br>";

        echo "<br><br>// ---  makeFontItalic($col, $str)---<br> ";
        $json = json_decode($str);
//$objPayable->getFont()->setItalic(true);  
//$objPayable->getFont()->setColor( new PHPExcel_Style_Color( PHPExcel_Style_Color::COLOR_DARKGREEN ) );  


        foreach ($json as $key => $val) {
//            echo $key;
//            print_r($val);
            foreach ($val as $item) {
//                 echo $item;
                $str = "\$objPHPExcel->getActiveSheet()->getStyle('$col$item')->getFont()->setItalic(true)->setBold(true);<br>";
                $str .= "\$objPHPExcel->getActiveSheet()->getStyle('$col$item')->getFont()->setColor( new PHPExcel_Style_Color('$key'));<br>";
                
             //   $str.=" ->getStartColor()->setARGB('$key');<br>";
                echo $str;
            }
        }
    }

}
