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
            var_dump($_POST['user']);die;
        }
    }

    $exec = new Excel();
    $build = $exec->buildReport();


?>