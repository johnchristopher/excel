<?php namespace JC;

use PHPExcel;
use PHPExcel_Cell;
use PHPExcel_Style_Border;
use PHPExcel_Style_Fill;
use PHPExcel_IOFactory;

class Excel
{
    public static function sendExcelFile($objPHPExcel, $filename, $parameters)
    {
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $parameters['writer']);

        header('Content-type: ' . $parameters['content-type']);
        header('Content-Disposition: attachment; filename="' . $filename . '.' . $parameters['extension'] . '"');
        header('Cache-Control: max-age=0');

        $objWriter->save('php://output');
    }

    public static function getPrettyObjPHPExcelFromHashes($hashes, $metadata)
    {
        $objPHPExcel = new PHPExcel();
        $objPHPExcel->getProperties()
                    ->setCreator($metadata['creator'])
                    ->setLastModifiedBy($metadata['lastmodifiedby'])
                    ->setTitle($metadata['title'])
                    ->setSubject($metadata['subject'])
                    ->setCompany($metadata['company'])
                    ->setDescription($metadata['description']);
        $objPHPExcel->getActiveSheet()->setTitle($metadata['sheet_title']);

        $rows = Tools::getArrayFromHashesWithKeysAsHeaders($hashes);
        $objPHPExcel->getActiveSheet()->fromArray($rows, null, 'A1');

        $last_column = PHPExcel_Cell::stringFromColumnIndex(count($rows[0]) - 1);
        $last_row = $objPHPExcel->getActiveSheet()->getHighestRow();

        for ($i = 0; $i < count($rows[0]); $i++) {
            $objPHPExcel->getActiveSheet()
                        ->getColumnDimension(PHPExcel_Cell::stringFromColumnIndex($i))
                        ->setAutoSize(true);
        }

        $header_border_style = array(
            'borders' => array(
                'bottom' => array(
                    'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                    'color' => array(
                        'argb' => '000000'))));
        $footer_border_style = array(
            'borders' => array(
                'bottom' => array(
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array(
                        'argb' => '000000'))));
        $objPHPExcel->getActiveSheet()
                    ->getStyle('A1:' . $last_column . '1')
                    ->applyFromArray($header_border_style);
        $objPHPExcel->getActiveSheet()
                    ->getStyle('A1:' . $last_column . '1')
                    ->getFont()
                    ->setBold(true);
        $objPHPExcel->getActiveSheet()
                    ->getStyle('A' . $last_row . ':' . $last_column . $last_row)
                    ->applyFromArray($footer_border_style);
        for ($row = 2; $row <= $last_row; $row++) {
            if ($row % 2 != 0) {
                $objPHPExcel->getActiveSheet()
                            ->getStyle('A' . $row . ':' . $last_column . $row)
                            ->applyFromArray(
                                array(
                                    'fill' => array(
                                        'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                        'color' => array(
                                        'rgb' => 'f2f2f2')))
                            );
            }
        }
        $objPHPExcel->getActiveSheet()->setSelectedCell('A1');

        return $objPHPExcel;
    }
}
