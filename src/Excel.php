<?php

namespace PFinal\Excel;

use DateTime;

/**
 * Excel操作类
 * @author  Zou Yiliang <it9981@gmail.com>
 * @since   1.0
 */
class Excel
{
    // //向下合并3行 $sheet->mergeCellsByColumnAndRow($col, $row + $rowOffset, $col, $row + $rowOffset + 2);

    /**
     * @return \PHPExcel
     */
    public static function makePHPExcel()
    {
        //单元格缓存到PHP临时文件中
        $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '900MB');
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

        //实例化工作簿对象
        return new \PHPExcel();
    }

    /**
     * 保存或下载excel
     *
     * 文件保存到服务器，$saveName指定保存的文件名
     *
     * 直接下载时，弹出下载对话框:
     * $saveName = 'php://output'
     *
     * @param \PHPExcel $objPHPExcel
     * @param string $saveName
     * @throws \PHPExcel_Reader_Exception
     */
    public static function savePHPExcel(\PHPExcel $objPHPExcel, $saveName)
    {

        if (strtolower($saveName) == 'php://output') {
            header('Content-Type: application/octet-stream');
            header('Content-Disposition: attachment; filename=' . date('YmdHis') . '.xls');
        }

        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save($saveName);
    }

    /**
     * 导出到Excel
     *
     * 使用示例
     *
     * $data = array(
     *     array('id' => 1, 'name' => 'Jack', 'age' => 18),
     *     array('id' => 2, 'name' => 'Mary', 'age' => 20),
     *     array('id' => 3, 'name' => 'Ethan', 'age' => 34),
     * );
     * $map = array(
     *     'title'=>array('id' => '编号',
     *          'name' => '姓名',
     *          'age' => '年龄',
     *      )
     * );
     * $file = 'user' . date('Y-m-d');
     * $excel = new \PHPExcel\Excel();
     * $excel->exportExcel($data, $map, $file, '用户信息');
     *
     * @param array $data 需要导出的数据
     * @param array $map 格题、数据格式、数字样式
     *      array(
     *              'title' =>array('id'=>'编号','name'=>'姓名'),
     *              'dataType' =>array('name'=>\PHPExcel_Cell_DataType::TYPE_STRING),
     *              'numberFormat' =>array('created_at' => \PHPExcel_Style_NumberFormat::FORMAT_DATE_DDMMYYYY),
     *      )
     * @param string $filename 下载显示的默认文件名
     * @param string $title 工作表名称
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */
    public static function exportExcelV1($data, $map = array(), $filename = '', $title = 'Worksheet')
    {
        if (!is_array($data)) {
            return;
        }
        if (count($data) < 1) {
            return;
        }

        //单元格缓存到PHP临时文件中
        $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '900MB');
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

        //实例化工作簿对象
        $objPHPExcel = new \PHPExcel();
        //获取活动工作表
        $objActSheet = $objPHPExcel->getActiveSheet();
        //设置工作表的标题
        $objActSheet->setTitle($title);

        //第一行为标题
        $col = 0;

        foreach ($data[0] as $key => $value) {
            if (isset($key, $map['title'][$key])) {
                $title = $map['title'][$key];
            } else {
                $title = $key;
            }
            $objActSheet->getCellByColumnAndRow($col, 1)->setValue($title);

            $col++;
        }

        //第2行开始是内容
        $row = 2;
        foreach ($data as $v) {

            //第一列序号
            //$objActSheet->getCellByColumnAndRow(0,$row)->setValue($row-1);

            $col = 0;
            foreach ($v as $key => $value) {

                if (isset($key, $map['dataType'][$key])) {
                    $pDataType = $map['dataType'][$key];


                    $objActSheet->getCellByColumnAndRow($col, $row)
                        ->setValueExplicit($value, $pDataType);

                } else {
                    $objActSheet->getCellByColumnAndRow($col, $row)
                        ->setValue($value);
                }

                if (isset($key, $map['numberFormat'][$key])) {
                    $numberFormat = $map['numberFormat'][$key];
                } else {
                    $numberFormat = \PHPExcel_Style_NumberFormat::FORMAT_GENERAL;
                }

                $objActSheet->getStyleByColumnAndRow($col, $row)
                    ->getNumberFormat()
                    ->setFormatCode($numberFormat);

                $col++;
            }
            $row++;
        }

        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

        if (empty($filename)) {
            $filename = date('YmdHis');
        }

        if (strtolower(substr($filename, -4)) != '.xls') {
            $filename .= '.xls';
        }

        //弹出下载对话框
        header('Content-Type: application/octet-stream');
        header('Content-Disposition: attachment; filename=' . $filename);

        $objWriter->save('php://output');
    }

    /**
     * 导出到Excel文件
     *
     * Office 2007+ xlsx format
     * supports writing huge 100K+ row spreadsheets
     *
     * 使用示例
     *
     * $data = array(
     *     array('id' => 1, 'name' => 'Jack', 'age' => 18),
     *     array('id' => 2, 'name' => 'Mary', 'age' => 20),
     *     array('id' => 3, 'name' => 'Ethan', 'age' => 34),
     * );
     *
     * $map = array(
     *     'title'=>array('id' => '编号',
     *          'name' => '姓名',
     *          'age' => '年龄',
     *      )
     * );
     *
     * $file = 'user' . date('Y-m-d');
     * $excel = new \PHPExcel\Excel();
     * $excel->exportExcel($data, $map, $file, '用户信息');
     *
     * @param $data
     * @param array $map
     * @param string $filename 保存的文件名 如果为空，将在临时目录生成随机文件名
     * @param string $workSheetName
     * @return string 返回文件名
     */
    public static function toExcelFile($data, $map = array(), $filename = '', $workSheetName = 'Worksheet')
    {
        /*
        $map = array(
            'title' => array('product_id' => '编号', 'created' => '时间', 'quantity' => '数量', 'amount' => '单价','description'=>'描述'),
            'simpleFormat' => array(
                'created' => 'date',
                'product_id' => 'integer',
                'quantity' => '#,##0',
                'amount' => 'price',
                'description' => 'string',
            ),
        );

        $data = array(
            array('created' => '2015-01-01', 'product_id' => 873, 'quantity' => 1, 'amount' => '44.00', 'description' => 'misc'),
            array('created' => '2015-01-12', 'product_id' => 324, 'quantity' => 2, 'amount' => '88.00', 'description' => 'none'),
        );//*/

        if (!isset($map['title'])) {
            if (count($data) > 0 && isset($data[0])) {
                $map['title'] = array_combine(array_keys($data[0]), array_keys($data[0]));
            } else {
                $map['title'] = array();
            }
        }

        $header = array();
        foreach ($map['title'] as $key => $val) {
            if (isset($map['simpleFormat'][$key])) {
                $header[$val] = $map['simpleFormat'][$key];
            } else {
                $header[$val] = 'GENERAL';
            }
        }

        $writer = new \XLSXWriter();
        //$writer->writeSheet($data, 'Sheet1', $header);

        $writer->writeSheetHeader($workSheetName, $header);
        foreach ($data as $row) {
            $temp = array();
            foreach ($map['title'] as $key => $value) {
                if (isset($row[$key])) {
                    $temp[] = $row[$key];
                } else {
                    $temp[] = '';
                }
            }

            $writer->writeSheetRow($workSheetName, $temp);
        }

        if (empty($filename)) {
            $filename = tempnam(sys_get_temp_dir(), 'excel');
        }
        /**
        * 此处不应有次判断,否则将造成读写不一致。无法正常导出excel文件
        if (strtolower(substr($filename, -5)) != '.xlsx') {
            $filename .= '.xlsx';
        }
        */
        $writer->writeToFile($filename);

        return $filename;
    }

    /**
     * 导出到Excel，浏览器直接下载
     *
     * Office 2007+ xlsx format
     * supports writing huge 100K+ row spreadsheets
     *
     * 使用示例
     *
     * $data = array(
     *     array('id' => 1, 'name' => 'Jack', 'age' => 18),
     *     array('id' => 2, 'name' => 'Mary', 'age' => 20),
     *     array('id' => 3, 'name' => 'Ethan', 'age' => 34),
     * );
     *
     * $map = array(
     *     'title'=>array('id' => '编号',
     *          'name' => '姓名',
     *          'age' => '年龄',
     *      )
     * );
     *
     * $file = 'user' . date('Y-m-d');
     * $excel = new \PHPExcel\Excel();
     * $excel->exportExcel($data, $map, $file, '用户信息');
     *
     * @param $data
     * @param array $map
     * @param string $filename
     * @param string $workSheetName
     */
    public static function exportExcel($data, $map = array(), $filename = '', $workSheetName = 'Worksheet')
    {
        //弹出下载对话框
        header('Content-Type: application/octet-stream');
        header('Content-Disposition: attachment; filename=' . $filename);

        $tempFile = tempnam(sys_get_temp_dir(), 'excel');

        static::toExcelFile($data, $map, $tempFile, $workSheetName);

        readfile($tempFile);
        unlink($tempFile);
    }

    /**
     * 读取Excel文件数据 (或cvs文件)
     *
     * 使用示例
     *
     * $map = array(
     *     'id' => '编号',
     *     'name' => '姓名',
     *     'age' => '年龄',
     *  );
     * $excel = new \PHPExcel\Excel();
     * $data = $excel->readExcelFile('./user2014-11-08.xls', $map);
     *
     * @param $file
     * @param array $map
     * @param int $titleRow 标题在第几行
     * @param int $beginColumn 从行几列开始
     * @param string $method 获取数据方式。 支持公式计算:calculated、格式化后返回:formatted、 默认为空:原样返回单元格数据
     * @param string $encoding 文件编码 UTF-8、GBK等。 导入csv文件时默认为UTF-8，如果你导入的csv文件是其它编码格式例如"GBK"，请传入此参数为"GBK"
     * @return array
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */
    public static function readExcelFile($file, $map = array(), $titleRow = 1, $beginColumn = 1, $method = '', $encoding = null)
    {
        //单元格缓存到PHP临时文件中
        $cacheMethod = \PHPExcel_CachedObjectStorageFactory:: cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '900MB');
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

        //$excelReader = \PHPExcel_IOFactory::createReader('Excel5');
        $excelReader = \PHPExcel_IOFactory::createReaderForFile($file);

        if ($encoding !== null && method_exists($excelReader, 'setInputEncoding')) {
            $excelReader->setInputEncoding($encoding);
        }

        //读取excel文件中的第一个工作表
        $phpExcel = $excelReader->load($file)->getSheet(0);

        //取得最大的行号
        $total_line = $phpExcel->getHighestRow();

        //取得最大的列号
        $total_column = $phpExcel->getHighestColumn();
        $highestColumnIndex = \PHPExcel_Cell::columnIndexFromString($total_column);

        //将列名与map对应
        $title = array();
        if ($titleRow > 0) {
            for ($cols = $beginColumn - 1; $cols < $highestColumnIndex; $cols++) {
                $val = $phpExcel->getCellByColumnAndRow($cols, $titleRow)->getValue();

                $field = array_search($val, $map);
                if ($field === false) {
                    $field = trim($val);
                }
                $title[] = $field;

            }
        }

        $data = array();
        $row = 0;
        for ($currentRow = $titleRow + 1; $currentRow <= $total_line; $currentRow++) {
            $i = 0;
            for ($cols = $beginColumn - 1; $cols < $highestColumnIndex; $cols++) {

                //单元格类型
                //PHPExcel_Cell_DataType::TYPE_STRING
                //str、s、f、n、b、null、inlineStr、e
                //$dataType = $phpExcel->getCellByColumnAndRow($cols, $currentRow)->getDataType();
                //var_dump($dataType);

                //单元格数据
                switch ($method) {
                    case 'calculated':
                        $val = $phpExcel->getCellByColumnAndRow($cols, $currentRow)->getCalculatedValue();//支持计算
                        break;
                    case 'formatted':
                        $val = $phpExcel->getCellByColumnAndRow($cols, $currentRow)->getFormattedValue(); //格式化的内容
                        break;
                    default:
                        $val = $phpExcel->getCellByColumnAndRow($cols, $currentRow)->getValue();//原始值
                        break;
                }

                $field = isset($title[$i]) ? $title[$i] : $i;
                $data[$row][$field] = trim($val);
                $i++;
            }
            $row++;
        }
        return $data;
    }

    /**
     * 将Excel文件以Html格式返回
     * @param $file
     * @return string
     */
    public static function toHTML($file)
    {
        //单元格缓存到PHP临时文件中
        $cacheMethod = \PHPExcel_CachedObjectStorageFactory:: cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '900MB');
        \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

        //$excelReader = PHPExcel_IOFactory::createReader('Excel5');
        $excelReader = \PHPExcel_IOFactory::createReaderForFile($file);
        //读取excel文件中的第一个工作表
        $phpExcel = $excelReader->load($file);
        $excelHTML = new Excel_HTML($phpExcel);
        return $excelHTML->ToHtml();
    }

    /**
     * 转换Excel日期时间类型
     * @param int|float $excelTime
     * @param string $format
     * @return string
     */
    public static function convertTime($excelTime, $format = 'Y-m-d H:i:s')
    {
        //EXCEL的date类型存的是从1900-1-1日开始算的，单位是天 整数或浮点数
        //EXCEL中 1970-1-1 代表的数字是25569
        //PHP 的时间函数是从1970-1-1日开始计算的，单位是秒

        if (!preg_match('/^\d+(\.\d+)?$/', $excelTime)) {
            return $excelTime;
        }

        $timestamp = intval(($excelTime - 25569) * 24 * 60 * 60);// 转化为PHP的时间戳

        $date = new DateTime();
        $date->setTimezone(new \DateTimeZone('UTC')); // Excel中已做时区处理
        $date->setTimestamp($timestamp);

        return $date->format($format);
    }

}
