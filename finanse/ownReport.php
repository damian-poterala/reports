<?php 
    require_once '../Classes/PHPExcel.php';
    require ('../../Cezar/config.php');

    class Excel extends Cezar 
    {
        function __construct()
        {
            $this->connect = $this->pdo_mysql();
        }

        function buildReport()
        {
            $year         = $_GET['year'];
            $month        = $_GET['month'];
            $numberBill   = $_GET['numberBill'];
            $categoryBill = $_GET['categoryBill'];
            $payment      = $_GET['payment'];
            $amountBill   = $_GET['amountBill'];
            $dateBill     = $_GET['dateBill'];
            $addDateBill  = $_GET['addDateBill'];
            $description  = $_GET['description'];
            $averageBill  = $_GET['averageBill'];

            var_dump($year, $month, $numberBill, $categoryBill, $payment, $amountBill, $dateBill, $addDateBill, $description, $averageBill);die;
        }
    }

    $exec = new Excel();
    $build = $exec->buildReport();


?>