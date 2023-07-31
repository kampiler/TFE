<?php
  require_once('tfe.php');

  
  $ox=new TMyExcel();

  //первый лист
  $ox->AddSheet("Лист1");
  $ox->AddAutoFilter();
  $ox->AddHead('Head 1',1);
  $ox->AddHead('Head 2',2);
  $ox->AddHead('Head 3',3);
  $ox->AddTimeStamp();

  $ox->AddTitle('Title 1');
  $ox->AddTitle('Title 2');
  $ox->AddTitle('Title 3');
  $ox->AddTitle('Title 4','Title 5','Title 6');

  $ox->AddCell("Hello");
  $ox->AddWidth(100);
  $ox->AddCell();
  $ox->AddWidth(200);
  $ox->AddCell("я купил жене и дочке\nафриканские\nчулочки они стоят три рубля очень дороги для меня\x1B");
  $ox->AddRow();
  $ox->AddCell(1,2,3,44,55.11,123.45,-678.90);

  $ox->AddRow();
  $ox->AddCell(88);
  $ox->AddCell(77.2222);
  $ox->AddCell(66/7);
  $ox->AddCell(55);
  $ox->AddCell(44);
  $ox->AddCell(33);
  $ox->AddCell(22);
  $ox->AddCell(11);
  $ox->AddRow();
  $ox->AddCell('11.11.2018');
  $ox->AddCell('2018-10-10');
  $ox->AddCell(date('Y-m-d H:i:s'));
  $ox->AddCell('17:23');

  $ox->AddTotal('Итого:');
  $ox->AddTotal(77,88,7878.113,66,1.23,4.56,1.20);
  $ox->AddTotal(66);

  // второй лист
  $ox->AddSheet('имя листа');
  $ox->AddTimeStamp();
  $ox->AddCell("ПРИВЕТ");
  $ox->AddCell("МИР");
  $ox->AddRow();
  $ox->AddCell();
  $ox->AddCell(1);
  $ox->AddCell(2);
  $ox->AddCell(3);
  $ox->AddCell(4);
  $ox->AddRow();
  $ox->AddCell(88);
  $ox->AddCell(77);
  $ox->AddCell(66);
  $ox->AddCell(55);
  $ox->AddCell(44.55);
  $ox->AddCell(33.3);
  $ox->AddCell(22.2222);
  $ox->AddCell(11.11);
  $ox->AddRow();

  // третий лист
  $ox->AddSheet();
  $ox->AddTimeStamp();
  $ox->AddCell('Таблица умножения','','',' (пифагора)');
  $ox->AddRow();

  for($i=1;$i<10;$i++)
    {
     for($j=1;$j<10;$j++)
       {
        $ox->AddCell($i*$j);
       }
     $ox->AddRow();
    }

  echo $ox->Go(2);
?>