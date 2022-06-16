<?php

require_once 'Excel.php';

class App
{
    /**
     * @var string �������� ����.
     */
    private $date;

    /**
     * @var string ���� ����� ����� (d.m.Y).
     */
    private $datePoint;

    /**
     * @var string ��� �������� ����.
     */
    private $year;

    /**
     * @var string ����� �������� ����.
     */
    private $month;

    /**
     * @var string ���� �������� ����.
     */
    private $day;

    /**
     * ������ ���������.
     * @return false
     */
    public function run()
    {
        if ($this->setDate()) {
            $fileParse = "E:\\Diskor_new\\TEP\\tep_prost_loc_5hour\\{$this->datePoint}_tep_prost_loc_5hour.xlsx";

            Log::Info("��������� ����� ��� �������� �����.");
            $xls = Excel::openExcel($fileParse);
            try {
                $dataParse = Excel::getDataExcel($xls);
            } catch (Exception $e) {
                log::Error($e->getMessage());
                $dataParse = array();
            }

            /**
             * $dataParse[5] - ������ ������� ����� 5 �����.
             * $dataParse[6] - ������ ������� ����� 6 �����.
             */

            $fileTemplate = "template.xlsx";
            $fileSave = "E:\\Diskor\\doc_boss\\gvc\\prost_lok_5hours.xlsx";
            Log::Info("�������� ������������ ������� � ������� ����� 5 �����");
            $xls = Excel::openExcel($fileTemplate);
            $xls = Excel::formattedExcel5hours($xls, $dataParse[5], $this->datePoint);
            Excel::saveExcelDoc($xls, $fileSave);
            $fileSave = "E:\\Diskor_new\\archiv\\$this->year\\$this->month\\$this->day\\prost_lok_5hours.xlsx";
            Excel::saveExcelDoc($xls, $fileSave);

            Log::Info("�������� ������������ ������� � ������� ����� 6 �����");
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
     * ������� ���� ��� ���������.
     */
    private function setDate()
    {
        $date = date('Y-m-d', time());

        if (isset($_REQUEST['date']) && $_REQUEST['date'] != '') {
            $date = $_REQUEST['date'];

            Log::Info("���� ��� ��: $date");

        } elseif (isset($_SERVER['argv'][1])) {
            Log::Info("���������� ������ �� ����� ������� �����");
            $date = $_SERVER['argv'][1];

            Log::Info("���� ��� ��: $date");
        } else {
            echo "�� ������� ���� ��� ������ � �������! ���������� ������ ���������...";
            Log::Error("�� ������� ���� ��� ������ � �������! ���������� ������ ���������...");
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