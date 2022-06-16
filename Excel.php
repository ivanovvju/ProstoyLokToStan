<?php

require_once 'D:\wwwnew\Classes\PHPExcel.php';

class Excel
{
    public function __construct($date)
    {
//        $dateObj = new DateTime($date);
//        $dateObjPr = $dateObj->modify('0 day');
//        $this->datePoint = $dateObjPr->format('d.m.Y');
//
//        $this->pathDoc = "{$this->datePoint}_tep_prost_loc_5hour.xlsx";

    }

    /**
     * ������� ������ Excel.
     * @param $path string ���� �� �����.
     * @return false|PHPExcel
     */
    public static function openExcel($path)
    {
        if (!file_exists($path)) {
            Log::Error("����������� ����: $path");
            return false;
        }

        try {
            echo "��������� excel-����. ($path)";
            $xls = PHPExcel_IOFactory::load($path);

            return $xls;
        } catch (PHPExcel_Reader_Exception $e) {
            Log::Error("��������� ������ �� ����� �������� excel-����� (�� ������): {$e->getMessage()}");
            return false;
        }
    }

    /**
     * ������ ������ �� excel.
     * @param $xls PHPExcel ������ Excel ��������� xls �� ������.
     * @return array|false ������ �������. ���� ����� ���, �� false.
     */
    public static function getDataExcel($xls)
    {
        if (!$xls) {
            throw new Exception("������ xls �������� ������!");
        }

        $xls->setActiveSheetIndex(0);
        $sheet = $xls->getActiveSheet();

        $data = $sheet->toArray();
        $parseData = array();

        echo "������ ����.";
        Log::Info("������ ����.");

        for ($row = 7; $row < count($data); $row++) {
            $times = explode(' ', $data[$row][14]);
            $hours = preg_replace('/[^,.0-9]/', '', $times[0]);

            if ($hours >= 5 && $data[$row][1] != 1.5) {
                for ($cell = 0; $cell < 15; $cell++) {
                    if ($cell == 8 || $cell == 12) {
                        $stringDateTime = $data[$row][$cell];
                        $format = 'n/j/y G:i';
                        $dateTime = DateTime::createFromFormat($format, $stringDateTime);
                        $data[$row][$cell] = $dateTime->format('d.m.Y H:i');
                    }
                    $parseData[5][$row][$cell] = $data[$row][$cell];
                    if ($hours >= 6) {
                        $parseData[6][$row][$cell] = $data[$row][$cell];
                    }
                }
            }
        }

        echo "��������� �������.";
        Log::Info("��������� �������.");

        return $parseData;
    }

    /**
     * ���������� excel-������� ������� - ������������ ������� � ��������� 5 �����.
     * @param $xls PHPExcel ������ ��������� excel.
     * @param $data array ������ � ������� ��� ���������� �������.
     * @param $datePoint string ���� (����� �����) ��� ������� � ������������ �������.
     * @return PHPExcel
     */
    public static function formattedExcel5hours($xls, $data, $datePoint)
    {
        $styleCell = array(
            'alignment' => array(
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            ),
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000'),
                ),
            ),
        );

        $xls->setActiveSheetIndex(0);
        $sheet = $xls->getActiveSheet();

        $title = iconv("cp1251", "utf-8", "������� ����������� �� ������� ����� 5 ����� (�� �������� �� �����������) �� $datePoint");
        $sheet->setCellValue('A2', $title);

        $row = 7;
        $sumSeconds = 0;

        if (!empty($data)) {
            foreach ($data as $dataRow) {

                $time = explode(" ", $dataRow[14]);
                $time = (int) $time[0] . ":" . (int) $time[1];
                $timeArr = explode(":", $time);

                if ((int) $dataRow[1] == 2) {
                    $seconds = ($timeArr[0] * 3600 + $timeArr[1] * 60) * 2;

                    $zero = new DateTime("@0");
                    $offset = new DateTime("@$seconds");
                    $diff = $zero->diff($offset);
                    $timeProstoy = sprintf("%02d:%02d", $diff->days * 24 + $diff->h, $diff->i);
                    $sumSeconds += $seconds;
                } else {
                    $seconds = $timeArr[0] * 3600 + $timeArr[1] * 60;
                    $sumSeconds += $seconds;
                    $timeProstoy = $time;
                }

                $sheet->setCellValueExplicitByColumnAndRow(0, $row, $row - 6, PHPExcel_Cell_DataType::TYPE_NUMERIC);
                $sheet->getStyleByColumnAndRow(0, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[1];
                $sheet->setCellValueExplicitByColumnAndRow(1, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(1, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[2];
                $sheet->setCellValueExplicitByColumnAndRow(2, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(2, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[3];
                $sheet->setCellValueExplicitByColumnAndRow(3, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(3, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[4];
                $sheet->setCellValueExplicitByColumnAndRow(4, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(4, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[7];
                $sheet->setCellValueExplicitByColumnAndRow(5, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(5, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[8];
                $sheet->setCellValueExplicitByColumnAndRow(6, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(6, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[12];
                $sheet->setCellValueExplicitByColumnAndRow(7, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(7, $row)->applyFromArray($styleCell);
                $dataCell = $time;
                $sheet->setCellValueExplicitByColumnAndRow(8, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(8, $row)->applyFromArray($styleCell);
                $dataCell = $timeProstoy;
                $sheet->setCellValueExplicitByColumnAndRow(9, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(9, $row)->applyFromArray($styleCell);

                $row++;
            }
            Log::Info("��������� �������� ������");
        } else {
            Log::Error("������ ��� ���������� ������� ���. ������� ����� ������!");
        }

        return $xls;
    }

    /**
     * ���������� excel-������� ������� - ������������ ������� � ��������� 6 �����.
     * @param $xls PHPExcel ������ ��������� excel.
     * @param $data array ������ � ������� ��� ���������� �������.
     * @param $datePoint string ���� (����� �����) ��� ������� � ������������ �������.
     * @return PHPExcel
     */
    public static function formattedExcel6hours($xls, $data, $datePoint)
    {
        $styleCell = array(
            'alignment' => array(
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            ),
            'borders' => array(
                'allborders' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => '000000'),
                ),
            ),
        );

        $xls->setActiveSheetIndex(0);
        $sheet = $xls->getActiveSheet();

        $title = iconv("cp1251", "utf-8", "������� ����������� �� ������� ����� 6 ����� (�� �������� �� �����������) �� $datePoint");
        $sheet->setCellValue('A2', $title);

        $row = 7;

        $sumSeconds = 0;

        if (!empty($data)) {
            foreach ($data as $dataRow) {

                $time = explode(" ", $dataRow[14]);
                $time = (int) $time[0] . ":" . (int) $time[1];
                $timeArr = explode(":", $time);

                if ((int) $dataRow[1] == 2) {
                    $seconds = ($timeArr[0] * 3600 + $timeArr[1] * 60) * 2;

                    $zero = new DateTime("@0");
                    $offset = new DateTime("@$seconds");
                    $diff = $zero->diff($offset);
                    $timeProstoy = sprintf("%02d:%02d", $diff->days * 24 + $diff->h, $diff->i);
                    $sumSeconds += $seconds;
                } else {
                    $seconds = $timeArr[0] * 3600 + $timeArr[1] * 60;
                    $sumSeconds += $seconds;
                    $timeProstoy = $time;
                }

                $sheet->setCellValueExplicitByColumnAndRow(0, $row, $row - 6, PHPExcel_Cell_DataType::TYPE_NUMERIC);
                $sheet->getStyleByColumnAndRow(0, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[1];
                $sheet->setCellValueExplicitByColumnAndRow(1, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(1, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[2];
                $sheet->setCellValueExplicitByColumnAndRow(2, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(2, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[3];
                $sheet->setCellValueExplicitByColumnAndRow(3, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(3, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[4];
                $sheet->setCellValueExplicitByColumnAndRow(4, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(4, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[7];
                $sheet->setCellValueExplicitByColumnAndRow(5, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(5, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[8];
                $sheet->setCellValueExplicitByColumnAndRow(6, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(6, $row)->applyFromArray($styleCell);
                $dataCell = $dataRow[12];
                $sheet->setCellValueExplicitByColumnAndRow(7, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(7, $row)->applyFromArray($styleCell);
                $dataCell = $time;
                $sheet->setCellValueExplicitByColumnAndRow(8, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(8, $row)->applyFromArray($styleCell);
                $dataCell = $timeProstoy;
                $sheet->setCellValueExplicitByColumnAndRow(9, $row, $dataCell, PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->getStyleByColumnAndRow(9, $row)->applyFromArray($styleCell);

                $row++;
            }

            $sumProstoy = round($sumSeconds / 60 / 60, 1);
            $lok = round($sumProstoy / 24, 1);

            $finishLine = "�������� � �������� ����������� ����� 6 ����� " . ($row - 7) . " ����� �� $sumProstoy ����������-����� �� ������������ ($lok ����������)";

            $sheet->mergeCellsByColumnAndRow(0, $row, 9, $row);
            $sheet->setCellValueByColumnAndRow(0, $row, iconv("cp1251", "utf-8", $finishLine), PHPExcel_Cell_DataType::TYPE_STRING);
            $sheet->getStyleByColumnAndRow(0, $row)->applyFromArray($styleCell);
            $sheet->getStyleByColumnAndRow(1, $row)->applyFromArray($styleCell);
            $sheet->getStyleByColumnAndRow(2, $row)->applyFromArray($styleCell);
            $sheet->getStyleByColumnAndRow(3, $row)->applyFromArray($styleCell);
            $sheet->getStyleByColumnAndRow(4, $row)->applyFromArray($styleCell);
            $sheet->getStyleByColumnAndRow(5, $row)->applyFromArray($styleCell);
            $sheet->getStyleByColumnAndRow(6, $row)->applyFromArray($styleCell);
            $sheet->getStyleByColumnAndRow(7, $row)->applyFromArray($styleCell);
            $sheet->getStyleByColumnAndRow(8, $row)->applyFromArray($styleCell);
            $sheet->getStyleByColumnAndRow(9, $row)->applyFromArray($styleCell);

            Log::Info("��������� �������� ������");
        } else {
            Log::Error("������ ��� ���������� ������� ���. ������� ����� ������!");
        }

        return $xls;
    }

    /**
     * ��������� excel-�������.
     * @param $xls PHPExcel ������ ��������� excel.
     * @param $path string ���� ���������� �������..
     * @return bool false ��� ������������� ������.
     */
    public static function saveExcelDoc($xls, $path)
    {
        try {
            $objWriter = new PHPExcel_Writer_Excel2007($xls);
            $objWriter->save($path);

            echo "������� ��������� �������.";
            Log::Info("������� ��������� ������� �: $path.");
        } catch (PHPExcel_Writer_Exception $e) {
            echo "��������� ������ �� ����� ���������� ������� �������: {$e->getMessage()}";
            Log::Error("��������� ������ �� ����� ���������� ������� �: $path: {$e->getMessage()}");
            return false;
        }

        return true;
    }
}