<?php
    require_once '../Classes/PHPExcel.php';
    require ('../../Cezar/config.php');

    class Excel extends Cezar
    {
        function __construct()
        {
            $this->connect = $this->pdo_mysql();
        }

        function list()
        {
            $sql = $this->connect->prepare("SELECT 	bill.id_rachunku,
                                                    bill.id_wyplaty,
                                                    year.rok,
                                                    CONCAT(year.rok, '-', payment.id_miesiaca) AS wyplata,
                                                    category.kat_nazwa_pl,
                                                    bill.kwota_rachunku,
                                                    bill.data_rachunku,
                                                    bill.data_dodania_rachunku,
                                                    type.nazwa_platnosci,
                                                    bill.dodatkowy_opis,
                                                    bill.numer_rachunku
                                            FROM FINANSE_rachunki AS bill
                                            LEFT JOIN finanse_wyplaty AS payment
                                                ON payment.id_wyplaty = bill.id_wyplaty
                                            LEFT JOIN FINANSE_rok AS year 
                                                ON payment.id_roku = year.id_roku
                                            LEFT JOIN FINANSE_kategorie_rachunkow AS category
                                                ON bill.id_kat_rachunku = category.id_kat_rachunkow
                                            LEFT JOIN FINANSE_typ_platnosci AS type 
                                                ON bill.id_platnosci = type.id_typ_platnosci
                                            LEFT JOIN konta AS user 
                                                ON bill.id_user = user.id_user
                                            WHERE bill.id_user = 1
                                            ORDER BY bill.data_rachunku DESC");

            $sql->execute();
            $result = $sql->fetchAll(PDO::FETCH_ASSOC);

            return $result;
        }
    }

    $exec = new Excel();
    $list = $exec->list();

    $objPHPExcel = new PHPExcel();

    $objPHPExcel->createSheet(0);
    $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getActiveSheet()->setTitle('Example');
    $objPHPExcel->getProperties()->setTitle('Example');

    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'ROK');
    $objPHPExcel->getActiveSheet()->setCellValue('B1', 'WYPŁATA');
    $objPHPExcel->getActiveSheet()->setCellValue('C1', 'KATEGORIA ZAKUPÓW');
    $objPHPExcel->getActiveSheet()->setCellValue('D1', 'KWOTA RACHUNKU');
    $objPHPExcel->getActiveSheet()->setCellValue('E1', 'DATA RACHUNKU');
    $objPHPExcel->getActiveSheet()->setCellValue('F1', 'DATA DODANIA RACHUNKU');
    $objPHPExcel->getActiveSheet()->setCellValue('G1', 'PŁATNOŚĆ');
    $objPHPExcel->getActiveSheet()->setCellValue('H1', 'OPIS ZAKUPÓW');
    $objPHPExcel->getActiveSheet()->setCellValue('I1', 'NUMER RACHUNKU');

    $objPHPExcel->getActiveSheet()->getStyle('A1:I1')->getFont()->setBold(true);
    $objPHPExcel->getActiveSheet()->freezePane('A2');

    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
    $objPHPExcel->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);

    $objPHPExcel->getActiveSheet()->getStyle('A1:I1')->applyFromArray(array(
       'borders' => array(
           'allborders' => array(
               'style' => PHPExcel_Style_Border::BORDER_THIN
           )
       ) 
    ));

    $objPHPExcel->getActiveSheet()->getStyle('A1:I1')->applyFromArray(array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
        )
    ));
    
    $objPHPExcel->getActiveSheet()->getStyle('A:I')->applyFromArray(array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
        ) 
    ));

    $objPHPExcel->getActiveSheet()->getStyle('A1:I1')->getFill()->applyFromArray(array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'startcolor' => array(
            'rgb' => 'FFAA00'
        )
    ));

    $row = 2;

    foreach($list as $key => $val)
    {
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0, $row, $val['rok']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1, $row, $val['wyplata']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(2, $row, $val['kat_nazwa_pl']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(3, $row, $val['kwota_rachunku']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(4, $row, $val['data_rachunku']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(5, $row, $val['data_dodania_rachunku']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(6, $row, $val['nazwa_platnosci']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(7, $row, $val['dodatkowy_opis']);
        $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(8, $row, $val['numer_rachunku']);

        $row++;
    }

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="item_list.xlsx"');
    header('Cache-Control: max-age=0');

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');   
    ob_end_clean();
    $objWriter->save('php://output');
?>