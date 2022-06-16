<?php

require_once 'Excel.php';

class App
{
    /**
     * @var string Отчетная дата.
     */
    private $date;

    /**
     * @var string Дата через точку (d.m.Y).
     */
    private $datePoint;

    /**
     * @var string Год отчетной даты.
     */
    private $year;

    /**
     * @var string Месяц отчетной даты.
     */
    private $month;

    /**
     * @var string День отчетной даты.
     */
    private $day;

    /**
     * Запуск программы.
     * @return false
     */
    public function run()
    {
        if ($this->setDate()) {
            $fileParse = "E:\\Diskor_new\\TEP\\tep_prost_loc_5hour\\{$this->datePoint}_tep_prost_loc_5hour.xlsx";

            Log::Info("Запускаем метод для парсинга файла.");
            $xls = Excel::openExcel($fileParse);
            try {
                $dataParse = Excel::getDataExcel($xls);
            } catch (Exception $e) {
                log::Error($e->getMessage());
                $dataParse = array();
            }

            /**
             * $dataParse[5] - данные простоя более 5 часов.
             * $dataParse[6] - данные простоя более 6 часов.
             */

            $fileTemplate = "template.xlsx";
            $fileSave = "E:\\Diskor\\doc_boss\\gvc\\prost_lok_5hours.xlsx";
            Log::Info("Начинаем формирвоание справки с данными более 5 часов");
            $xls = Excel::openExcel($fileTemplate);
            $xls = Excel::formattedExcel5hours($xls, $dataParse[5], $this->datePoint);
            Excel::saveExcelDoc($xls, $fileSave);
            $fileSave = "E:\\Diskor_new\\archiv\\$this->year\\$this->month\\$this->day\\prost_lok_5hours.xlsx";
            Excel::saveExcelDoc($xls, $fileSave);

            Log::Info("Начинаем формирвоание справки с данными более 6 часов");
            $xls = Excel::openExcel($fileTemplate);
            $xls = Excel::formattedExcel6hours($xls, $dataParse[6], $this->datePoint);
            $fileSave = "E:\\Diskor\\doc_boss\\gvc\\prost_lok_6hours.xlsx";
            Excel::saveExcelDoc($xls, $fileSave);
            $fileSave = "E:\\Diskor_new\\archiv\\$this->year\\$this->month\\$this->day\\prost_lok_6hours.xlsx";
            Excel::saveExcelDoc($xls, $fileSave);
        } else {
            return false;
        }

        return true;
    }


    /**
     * Зададим даты для программы.
     */
    private function setDate()
    {
        $date = date('Y-m-d', time());

        if (isset($_REQUEST['date']) && $_REQUEST['date'] != '') {
            $date = $_REQUEST['date'];

            Log::Info("Дата для ПО: $date");

        } elseif (isset($_SERVER['argv'][1])) {
            Log::Info("Произведен запуск ПО через шедулер задач");
            $date = $_SERVER['argv'][1];

            Log::Info("Дата для ПО: $date");
        } else {
            echo "Не указана дата для работы с файлами! Прекращаем работу программы...";
            Log::Error("Не указана дата для работы с файлами! Прекращаем работу программы...");
            return false;
        }

        $dateObj = new DateTime($date);
        $dateObjPr = $dateObj->modify('0 day');
        $this->datePoint = $dateObjPr->format('d.m.Y');
        $this->date = $dateObjPr->format('Y-m-d');
        $this->year = $dateObjPr->format('Y');
        $this->month = $dateObjPr->format('m');
        $this->day = $dateObjPr->format('d');

        return true;
    }
}