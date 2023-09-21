<?php
  error_reporting(E_ALL);
  date_default_timezone_set("Asia/Baghdad");

  function xlStr($s)
    {
     if(empty($s)) return '';
     $r=$s;
     $r=str_replace('<','&lt;',$r);
     $r=str_replace('>','&gt;',$r);
     $r=str_replace(chr(10),'&#10;',$r);
     $r=str_replace(chr(13),'&#13;',$r);
     $r=str_replace(chr(27),'&#27;',$r);
     return $r;
    }

  function is_money($money)
    {
     return preg_match("/^-?[0-9]+(?:\.[0-9]{1,2})?$/", $money);
    }

  function xlStyleCC($s)
    {
     $r='';
     if(preg_match("/^(\d{2}).(\d{2}).(\d{4})$/",$s,$m))
       {//короткая дата
        if(checkdate($m[2],$m[1],$m[3])) $r='cd';
       }
     elseif(preg_match("/^(\d{4}).(\d{2}).(\d{2})$/",$s,$m))
       {//короткая дата
        if(checkdate($m[2],$m[3],$m[1])) $r='cd';
       }
     elseif(preg_match("/^(\d{4}).(\d{2}).(\d{2}).(\d{2}).(\d{2}).(\d{2})$/",$s,$m))
       {//длинная дата
        if(checkdate($m[2],$m[1],$m[3])) $r='cg';
       }
     elseif(preg_match("/^(\d{2})\:(\d{2})$/",$s,$m)) $r='ct';//время
     elseif(preg_match("#^[0-9]+$#",$s,$m)) $r='cn';//целое число
     elseif(is_money($s)) $r='cm';//деньги
     elseif((strlen($s)>50)or(strpos($s,chr(13))>0)) $r='cw';//длинная строка
     else $r='c';//просто строка
     return $r;
    }

  function is_date($date, $format = 'Y-m-d H:i:s')
    {
     $d = DateTime::createFromFormat($format, $date);
     return $d && $d->format($format) == $date;
    }
  
  function xlDateTime($s)
  #checkdate(int $month, int $day, int $year)
    {
     $r='';
     if(preg_match("/^(\d{2}).(\d{2}).(\d{4})$/",$s,$m))
       {
        if(checkdate($m[2],$m[1],$m[3])) $r="$m[3]-$m[2]-$m[1]T00:00:00.000";
       }
     elseif(preg_match("/^(\d{4}).(\d{2}).(\d{2})$/",$s,$m))
       {
        if(checkdate($m[2],$m[3],$m[1])) $r="$m[1]-$m[2]-$m[3]T00:00:00.000";
       }
     elseif(preg_match("/^(\d{4}).(\d{2}).(\d{2}).(\d{2}).(\d{2}).(\d{2})$/",$s,$m))
       {
        if(checkdate($m[2],$m[3],$m[1])) $r="$m[1]-$m[2]-$m[3]T$m[4]:$m[5]:$m[6].000";
       }
     elseif(preg_match("/^(\d{2})\:(\d{2})$/",$s,$m))
       {
        $r="1899-12-31T$m[1]:$m[2]:00.000";
       }
     return $r;
    }

  function xlTypes($v)
    {
     $r='String';
     $s=gettype($v);
     if($s=='integer') $r='Number';//Integer
     elseif($s=='double') $r='Number';
     elseif($s=='string')
       {
        $r='String';
        if((preg_match('#^[0-9]+$#',$v,$m))) $r='Number';
        elseif(is_numeric($v)) $r='Number';
       }
     if((xlStyleCC($v)==='cd')or(xlStyleCC($v)==='cg')or(xlStyleCC($v)==='ct'))
       {
        $r='DateTime';
       }
     return($r);
    }

  class TMySheet//TFESheet
    {
     public $name;
     public $d = array();//data   ячейки
     public $s = array();//styles стиль конкретной ячейки
     public $ix=0,$iy=0;
     public $xRow=0;

     public $ixMax=0, $iyMax=0;

     public $head  = array();
     public $merge = array();
     public $width = array();
     public $title = array();
     public $total = array();

     public $RemainInp     = '';
     public $RemainOut     = '';
     public $PeriodStr     = '';

     public $useTimeStamp  = False;
     public $useTitleNum   = True;
     public $useAutoFilter = False;
     public $ixAutoFilter  = 0;//строка для автофильтра
     public $iyAutoFilter  = 0;//столбцов для автофильтра

     function __construct($n='')
       {
        if($n!='')
          {
           $this->name=$n;
          }
       }    

     function AddCell($data)
       {
        $this->d[$this->ix][$this->iy]=xlStr($data);
        $this->iy++;

        if($this->iy>$this->iyMax) $this->iyMax=$this->iy-1;
       }

     function AddRow()
       {
        $this->d[]=array();
        $this->ix++;
        $this->iy=0;
        $this->ixMax=$this->ix-1;
       }

     function __toString()
       {
        $r='';
        foreach($this->d as $i=>$arr)
          {
           if(gettype($arr)=='array')
             {
              $r.="\t<Row>\n";
              foreach($arr as $k=>$v)
                {
                 if(xlTypes($v)=='DateTime')
                   $r.="\t\t<Cell ss:StyleID='".xlStyleCC($v)."'><Data ss:Type='".xlTypes($v)."'>".xlDateTime($v)."</Data></Cell>\n";
                 elseif((xlTypes($v)=='Number')or(xlTypes($v)=='String'))
                   $r.="\t\t<Cell ss:StyleID='".xlStyleCC($v)."'><Data ss:Type='".xlTypes($v)."'>$v</Data></Cell>\n";
                 else
                   $r.="\t\t<Cell ss:StyleID='c'><Data ss:Type='".xlTypes($v)."'>$v</Data></Cell>\n";
                }
               $r.="\t</Row>\n";
             }
          }
        return $r;
       }
    }

  class TMyExcel//TFE
    {
     public $fileName,
            $ia, #индекс листа
            $types=array(),
            $sheets=array(),
            $styles=array();

     function __construct($aFirstSheetName='')
       {
        if($aFirstSheetName!='')
          {
           $this->ia=0;
           $this->fileName=$aFirstSheetName;
           $this->sheets[$this->ia]=new TMySheet($aFirstSheetName);
          }
       }    

     public function ActiveSheet():TMySheet
       {
        if(count($this->sheets)==0) return NULL;
        else return $sheets[$this->ia];
       } 

     function AddSheet($name='')
       {
        $this->ia++;
        if($name=='') $name='Лист'.$this->ia;
        if(strlen($name)>31) $name=substr($name,0,30);
        $this->sheets[$this->ia]=new TMySheet($name);
       }

     function AddCell(...$data)
       {
        if(func_num_args()==0) $this->sheets[$this->ia]->AddCell('');
        foreach($data as $d)
          $this->sheets[$this->ia]->AddCell($d);
       }

     function AddRow()
       {
        $this->sheets[$this->ia]->AddRow();
       }

     function AddTimeStamp()
       {
        $this->sheets[$this->ia]->useTimeStamp=True;
       }

     function AddPeriodStr($s)
       {
        $this->sheets[$this->ia]->PeriodStr=$s;
       }

     function AddWidth($w)
       {
        $this->sheets[$this->ia]->width[]=$w;
       }

     function AddHead($head, $merge=1)
       {
        $this->sheets[$this->ia]->head[]  = $head;
        $this->sheets[$this->ia]->merge[] = $merge-1;
       }

     function AddRemainInp($s)
       {
        $this->sheets[$this->ia]->RemainInp=$s;
       }

     function AddRemainOut($s)
       {
        $this->sheets[$this->ia]->RemainOut=$s;
       }

     function AddTitle(...$titles)
       {
        foreach($titles as $title)
          $this->sheets[$this->ia]->title[]=$title;
       }

     function AddTotal(...$totals)
       {
        foreach($totals as $total)
          $this->sheets[$this->ia]->total[]=$total;
       }

     function AddAutoFilter()
       {
        $this->sheets[$this->ia]->useAutoFilter=True;
       }


     function MakeFile()
       {
        $HeadColor ='Tan';#"#baacc7";
        $TitleColor='Aqua';#"#D6E5CB";
        $TotalColor="cyan";##eebef1
        $r='';
        // заголовки xml-файла
        $r.="<?xml version='1.0' encoding='utf-8'?>\n";
        $r.="<?mso-application progid='Excel.Sheet'?>\n";
        $r.="<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet'";
        $r.=" xmlns:o='urn:schemas-microsoft-com:office:office'";
        $r.=" xmlns:x='urn:schemas-microsoft-com:office:excel'";
        $r.=" xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'";
        $r.=" xmlns:html='http://www.w3.org/TR/REC-html40'>\n";
        $r.=" <DocumentProperties xmlns='urn:schemas-microsoft-com:office:office'>\n";
        $r.="  <Version>12.00</Version>\n";
        $r.=" </DocumentProperties>\n";
        $r.=" <ExcelWorkbook xmlns='urn:schemas-microsoft-com:office:excel'>\n";
        $r.="  <ProtectStructure>False</ProtectStructure>\n";
        $r.="  <ProtectWindows>False</ProtectWindows>\n";
        $r.=" </ExcelWorkbook>\n\n";
        $r.=" <Styles>\n";
        $r.="  <Style ss:ID='Default' ss:Name='Normal'>\n";
        $r.="   <Alignment ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="   <Borders/>\n";
        $r.="   <Font ss:FontName='Arial Cyr' x:CharSet='204'/>\n";
        $r.="   <Interior/>\n";
        $r.="   <NumberFormat/>\n";
        $r.="   <Protection/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='head'>\n";
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <Font ss:FontName='Arial Cyr' x:CharSet='204' ss:Bold='1'/>\n";
        $r.="   <Interior ss:Color='$HeadColor' ss:Pattern='Solid'/>\n";
        $r.="   <NumberFormat ss:Format='Standard'/>\n";
        $r.="   <Alignment ss:Horizontal='Center' ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='title'>\n";
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <Font ss:FontName='Arial Cyr' x:CharSet='204' ss:Bold='1'/>\n";
        $r.="   <Interior ss:Color='$TitleColor' ss:Pattern='Solid'/>\n";
        $r.="   <NumberFormat ss:Format='Standard'/>\n";
        $r.="   <Alignment ss:Horizontal='Center' ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='titleInt'>\n";
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <Font ss:FontName='Arial Cyr' x:CharSet='204' ss:Bold='1'/>\n";
        $r.="   <Interior ss:Color='$TitleColor' ss:Pattern='Solid'/>\n";
        $r.="   <NumberFormat ss:Format='0'/>\n";
        $r.="   <Alignment ss:Horizontal='Center' ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='total'>\n";
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <Font ss:FontName='Arial Cyr' x:CharSet='204' ss:Bold='1'/>\n";
        $r.="   <Interior ss:Color='$TotalColor' ss:Pattern='Solid'/>\n";
        $r.="   <NumberFormat ss:Format='Standard'/>\n";
        $r.="   <Alignment ss:Horizontal='Center' ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='c'>\n";#для всего остального - общий
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <Alignment ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='cw'>\n";#для длинных строк
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <Alignment ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='cd'>\n";#для коротких дат
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <NumberFormat ss:Format='Short Date'/>\n";
        $r.="   <Alignment ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='cg'>\n";#полная дата
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <NumberFormat ss:Format='General Date'/>\n";
        $r.="   <Alignment ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='ct'>\n";#время
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <NumberFormat ss:Format='h:mm;@'/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='cm'>\n";#деньги
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <NumberFormat ss:Format='#,##0.00_ ;\\-#,##0.00\\ '/>\n";
        $r.="   <Alignment ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="  </Style>\n";

        $r.="  <Style ss:ID='cn'>\n";#целое
        $r.="   <Borders>\n";
        $r.="    <Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="    <Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='1'/>\n";
        $r.="   </Borders>\n";
        $r.="   <NumberFormat ss:Format='0'/>\n";//
        $r.="   <Alignment ss:Vertical='Center' ss:WrapText='1'/>\n";
        $r.="  </Style>\n";

        $r.="
             <Style ss:ID='PeriodStr'>
               <Font ss:FontName='Verdana' x:CharSet='204' x:Family='Swiss' ss:Size='12' ss:Bold='1'/>
             </Style>\n";

        $r.="
             <Style ss:ID='TimeStamp'>
               <Font ss:FontName='Verdana' x:CharSet='204' x:Family='Swiss' ss:Size='7'/>
             </Style>\n";
        // вывели все стили
        $r.="</Styles>\n\n";
                
        foreach($this->sheets as $i=>$sh)
          {
           $r.=" <Worksheet ss:Name='".$sh->name."'>\n";#this->sheets[$i]
           $r.="  <Table>\n";

           //ширина столбцов
           if(count($sh->width)!==0)
             {
              foreach($sh->width as $w=>$width)
                $r.= "   <Column ss:AutoFitWidth='0' ss:Width='$width'/>\n";
             }
           
           // метка времени формирования отчета
           if($sh->useTimeStamp)
             {
              $sh->xRow++;
              $r.="   <Row>\n";
              $r.="    <Cell ss:StyleID='TimeStamp' ss:MergeAcross='$sh->iyMax'><Data ss:Type='String'>Отчет сформирован: ".date('d.m.Y - H:i:s')."</Data></Cell>\n";
              $r.="   </Row>\n";
             }

           // период
           if($sh->PeriodStr!='')
             {
              $sh->xRow++;
              $r.="   <Row>\n";
              $r.="    <Cell ss:StyleID='PeriodStr' ss:MergeAcross='$sh->iyMax'><Data ss:Type='String'>".$sh->PeriodStr."</Data></Cell>\n";
              $r.="   </Row>\n";
             }
              

           // входящий остаток +++
           if($sh->RemainInp!='')
             {
              $sh->xRow++;
              $r.="   <Row>\n";
              $r.="    <Cell ss:MergeAcross='$sh->iyMax'><Data ss:Type='String'>".$sh->RemainInp."</Data></Cell>\n";
              $r.="   </Row>\n";
             }
              
           //заголовки
           if(count($sh->head)!==0)
             {
              $sh->xRow++;
              $r.="   <Row>\n";
              foreach($sh->head as $h=>$head)
                {
                 $m=$sh->merge[$h];
                 $r.="    <Cell ss:StyleID='head' ss:MergeAcross='$m'><Data ss:Type='String'>$head</Data></Cell>\n";
                }
              $r.="   </Row>\n";
              $sh->ixAutoFilter = $sh->xRow;
              $sh->iyAutoFilter = count($sh->head);
             }
           
           //подзаголовки
           if(count($sh->title)!==0)
             {
              $sh->xRow++;
              $r.="   <Row>\n";
              foreach($sh->title as $t=>$title)
                 $r.="    <Cell ss:StyleID='title'><Data ss:Type='String'>$title</Data></Cell>\n";
              $r.="   </Row>\n";
              //номера подзаголовков
              if($sh->useTitleNum)
                {
                 $sh->xRow++;
                 $r.="   <Row>\n";
                 foreach($sh->title as $t=>$title)
                    $r.="    <Cell ss:StyleID='titleInt'><Data ss:Type='Number'>".($t+1)."</Data></Cell>\n";
                 $r.="   </Row>\n";
                }
              $sh->ixAutoFilter = $sh->xRow;
              $sh->iyAutoFilter = count($sh->title);
             }
           
           # тут выводим содержимое
           $r.=$sh;//toString

           //итоги
           if(count($sh->total)!==0)
             {
              $sh->xRow++;
              $r.="   <Row>\n";
              foreach($sh->total as $t=>$total)
                $r.="    <Cell ss:StyleID='total'><Data ss:Type='".xlTypes($total)."'>$total</Data></Cell>\n";
              $r.="   </Row>\n";
             }

           // исходящий остаток +++
           if($sh->RemainOut!='')
             {
              $sh->xRow++;
              $r.="   <Row>\n";
              $r.="    <Cell ss:MergeAcross='$sh->iyMax'><Data ss:Type='String'>".$sh->RemainOut."</Data></Cell>\n";
              $r.="   </Row>\n";
             }
              
           $r.="  </Table>\n";

           /*
           $r.="  <WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>\n";
           $r.="   <Selected/>\n";
           $r.="   <FreezePanes/>\n";
           $r.="   <FrozenNoSplit/>\n";
           $r.="   <SplitHorizontal>0</SplitHorizontal>\n";
           $r.="   <TopRowBottomPane>0</TopRowBottomPane>\n";
           $r.="   <ActivePane>0</ActivePane>\n";
           $r.="   <Panes>\n";
           $r.="    <Pane>\n";
           $r.="     <Number>3</Number>\n";
           $r.="    </Pane>\n";
           $r.="    <Pane>\n";
           $r.="     <Number>2</Number>\n";
           $r.="    </Pane>\n";
           $r.="   </Panes>\n";
           $r.="   <ProtectObjects>False</ProtectObjects>\n";
           $r.="   <ProtectScenarios>False</ProtectScenarios>\n";
           $r.="  </WorksheetOptions>\n";
           */

           if($sh->useAutoFilter)#'R4C1:R4C6'
             if($sh->ixAutoFilter>0)
               $r.="  <AutoFilter x:Range='R".$sh->ixAutoFilter."C1:R".$sh->ixAutoFilter."C".$sh->iyAutoFilter."' xmlns='urn:schemas-microsoft-com:office:excel'></AutoFilter>\n";

           $r.=" </Worksheet>\n\n";
          }

        $r.="</Workbook>\n";
        #echo "<textarea cols=50 rows=2>$r</textarea>";
        return $r;
       }

     function Go($p=0,$fn='123.xml')
       {
        $r=$this->MakeFile();
        if($p===1)
          {
           $r=htmlentities($r);
           $r=str_replace("\n","<br>\n",$r);
           $r=str_replace(" ","&nbsp;",$r);
          }
        elseif($p===2)
          {
           $size  = strlen($r);

           // Выводим HTTP-заголовки
           header ( "Expires: Mon, 1 Apr 1974 05:00:00 GMT" );
           header ( "Last-Modified: " . gmdate("D,d M YH:i:s") . " GMT" );
           header ( "Cache-Control: no-cache, must-revalidate" );
           header ( "Pragma: no-cache" );
           header ( "Content-type: application/vnd.ms-excel" );
           header ( "Content-Disposition: attachment; filename=$fn" );
           header ( "Content-Transfer-Encoding: binary");
           header ( 'Content-length: '.$size);
          }
        elseif($p===3)
          {
           if(file_put_contents($fn,$r)>0) $r=$fn;
          }
        return $r;
        //OpenFile();
       }
    }
?>