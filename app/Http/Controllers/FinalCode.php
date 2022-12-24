<?php

namespace App\Http\Controllers;
use PHPExcel_Style_Alignment;
use PHPExcel_Style_Fill;
use Illuminate\Http\Request;

class FinalCode extends Controller
{
    public function excel()
    {
        /**
        * Always refer to the package documentation for the latest example
        * @see http://phpexcel.codeplex.com/wikipage?title=Examples
        */
        // require __DIR__.'/vendor/autoload.php';

        // echo date('H:i:s') . " Create new PHPExcel object\n";
        $objPHPExcel = new \PHPExcel();
        echo date('H:i:s') . " Set properties\n";
        $objPHPExcel->getProperties()->setCreator("Maarten Balliauw");
        $objPHPExcel->getProperties()->setLastModifiedBy("Maarten Balliauw");
        $objPHPExcel->getProperties()->setTitle("Office 2007 XLSX Test Document");
        $objPHPExcel->getProperties()->setSubject("Office 2007 XLSX Test Document");
        $objPHPExcel->getProperties()->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.");

        $orders = [
          [
            'id'=>'1',
            'title'=>'Test Order',
            'createdat'=>'2022-12-23',
            'items'=>[
              [
                'id'=>'1',
                'qty'=>'15',
                'price'=>'100',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'2',
                'qty'=>'110',
                'price'=>'1000',
                'Photo' => public_path().'/image/image1.jpg',
              ],
              [
                'id'=>'3',
                'qty'=>'150',
                'price'=>'1000',
                'Photo' => public_path().'/image/image2.jpg',
              ]
            ]
          ],
          [
            'id'=>'2',
            'title'=>'Test Order 2assdfgh jkllgjeiogjd fgjdfgi odjgfi dougodgodyugiojdogudgodohj diogfdogjklhdngo jkdlhgiohd vdjklgh8d9 dklgd9ovb dk8e9djvbdm,eoiuhj',
            'createdat'=>'2022-11-23',
            'items'=>[
              [
                'id'=>'1',
                'qty'=>'25',
                'price'=>'2100',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'2',
                'qty'=>'210',
                'price'=>'21000',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'3',
                'qty'=>'250',
                'price'=>'31000',
                'Photo' => public_path().'/image/abcd.jpg',
              ]
            ]
          ],
          [
            'id'=>'3',
            'title'=>'Test Order 3 dfgh jkllgjeiogjd fgjdfgi odjgfi dougodgodyugiojdogudgodohj diogfdogjklhdngo jkdlhgiohd vdjklgh8d9 dklgd9ovb dk8e9djvbdm,eoiuhj',
            'createdat'=>'2022-10-23',
            'items'=>[
              [
                'id'=>'1',
                'qty'=>'45',
                'price'=>'4100',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'2',
                'qty'=>'510',
                'price'=>'51000',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'3',
                'qty'=>'650',
                'price'=>'61000',
                'Photo' => public_path().'/image/abcd.jpg',
              ]
            ]
          ],

          [
            'id'=>'4',
            'title'=>'Test Order 3 dfgh jkllgjeiogjd fgjdfgi odjgfi dougodgodyugiojdogudgodohj diogfdogjklhdngo jkdlhgiohd vdjklgh8d9 dklgd9ovb dk8e9djvbdm,eoiuhj',
            'createdat'=>'2022-10-23',
            'items'=>[
              [
                'id'=>'1',
                'qty'=>'45',
                'price'=>'4100',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'2',
                'qty'=>'510',
                'price'=>'51000',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'3',
                'qty'=>'650',
                'price'=>'61000',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'4',
                'qty'=>'650',
                'price'=>'61000',
                'Photo' => public_path().'/image/abcd.jpg',
              ]
            ]
          ],

          [
            'id'=>'5',
            'title'=>'Test Order 3 dfgh jkllgjeiogjd fgjdfgi odjgfi dougodgodyugiojdogudgodohj diogfdogjklhdngo jkdlhgiohd vdjklgh8d9 dklgd9ovb dk8e9djvbdm,eoiuhj',
            'createdat'=>'2022-10-23',
            'items'=>[
              [
                'id'=>'1',
                'qty'=>'45',
                'price'=>'4100',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'2',
                'qty'=>'510',
                'price'=>'51000',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'3',
                'qty'=>'650',
                'price'=>'61000',
                'Photo' => public_path().'/image/abcd.jpg',
              ]
            ]
          ],

          [
            'id'=>'6',
            'title'=>'Test Order 3 dfgh jkllgjeiogjd fgjdfgi odjgfi dougodgodyugiojdogudgodohj diogfdogjklhdngo jkdlhgiohd vdjklgh8d9 dklgd9ovb dk8e9djvbdm,eoiuhj',
            'createdat'=>'2022-10-23',
            'items'=>[
              [
                'id'=>'1',
                'qty'=>'45',
                'price'=>'4100',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'2',
                'qty'=>'510',
                'price'=>'51000',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'3',
                'qty'=>'650',
                'price'=>'61000',
                'Photo' => public_path().'/image/abcd.jpg',
              ],
              [
                'id'=>'4',
                'qty'=>'650',
                'price'=>'61000',
                'Photo' => public_path().'/image/abcd.jpg',
              ]
            ]
          ],
        ];


        $style = array(
          'alignment' => array(
              'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
          )
        );   


        $counter = 0;
        $sheet = $objPHPExcel->getActiveSheet();
        $ord_arr = 'A';
        $item_arr = 'B';
        $start_count = 1;
        $co_cunter = 1;
        $count = 1;
        $parentCounter = 1;
        $childCounter = 1;
        $aCounter = 1;
        $mergeTill = 0;
        $b_counter = 1;
        $c_counter = 1;
        $d_counter = 1;
        $e_counter = 1;
        $row_c = 1;
        $c_count = 4;
        $color_counter =1;
        $color =  "#FF0000";
        $col_count = 1; 
        foreach ($orders as $value) {
          
          $row = 'A'.$aCounter;
          $sheet->SetCellValue($row, $value['title']);
          $mergeTill += count($value['items']);
          $sheet->mergeCells('A'.$aCounter.':A'.$mergeTill); 
          $sheet->getStyle('A'.$aCounter.':A'.$mergeTill)->getAlignment()->setWrapText(true)->applyFromArray(
            array('vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER)
          );


          if($parentCounter % 2 == 0)
          {
            $objPHPExcel->getActiveSheet()->getStyle('A'.$aCounter.':A'.$mergeTill)->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()->setARGB('FFFF0000');   
            foreach ($value['items'] as $val) {
              $objPHPExcel->getActiveSheet()->getStyle('B'.$col_count.':E'.$col_count)->getFill()
              ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
              ->getStartColor()->setARGB('FFFF0000');
              $col_count++;
            }         
          }else
          {
            $objPHPExcel->getActiveSheet()->getStyle('A'.$aCounter.':A'.$mergeTill)->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()->setARGB('00FF00');
            $new_count = $col_count;
            foreach ($value['items'] as $val) {
              $objPHPExcel->getActiveSheet()->getStyle('B'.$new_count.':E'.$new_count)->getFill()
              ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
              ->getStartColor()->setARGB('00FF00');
              $new_count++;
              $col_count = $new_count;
            }
          }

          $aCounter = $mergeTill + 1;
          $count = 0;
          foreach ($value['items'] as $val) {   
            $row2 = 'B'.$b_counter;
            $row3 = 'C'.$c_counter;
            $row4 = 'D'.$d_counter;
            $row5 = 'E'.$e_counter;
            $sheet->SetCellValue($row2, $val['id']);
            $sheet->SetCellValue($row3, $val['qty']);
            $sheet->SetCellValue($row4, $val['price']);
            $objDrawing  = new \PHPExcel_Worksheet_Drawing();
            $objDrawing ->setPath($val['Photo']);    
            $objDrawing->setCoordinates($row5); 
            $objDrawing->setOffsetX(10); 
            $objDrawing->setOffsetY(10); 
            $objDrawing->setWidth(100); 
            $objDrawing->setHeight(150);
            $sheet->getRowDimension($row_c)->setRowHeight(150);
            $sheet->getColumnDimensionByColumn($c_count)->setWidth(50);
            $objDrawing->setWorksheet($sheet);
            $b_counter++;
            $c_counter++;
            $d_counter++;
            $e_counter++;
            $count++;
            $row_c++;
            $c_count++;
          }
          $parentCounter++;
          $color_counter++;
        }
        
        // Rename sheet
        echo date('H:i:s') . " Rename sheet\n";
        $objPHPExcel->getActiveSheet()->setTitle('Simple');

        // Save Excel 2007 file
        echo date('H:i:s') . " Write to Excel2007 format\n";
        $objWriter = new \PHPExcel_Writer_Excel2007($objPHPExcel);
        $objWriter->save(str_replace('.php', time().'.xlsx', __FILE__));
        // Echo done
        echo date('H:i:s') . " Done writing file.\r\n";
    }
}
