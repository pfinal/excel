<?php

include '../vendor/autoload.php';

use PFinal\Excel\Excel;

date_default_timezone_set('PRC');

$map = array(
    'title' => array(
        'id' => '编号',
        'name' => '姓名',
        'age' => '年龄',
    )
);


Excel::chunkExportCSV($map, './temp.csv', function ($writer) {

    $data = array(
        array('id' => 1, 'name' => 'Jack', 'age' => 18),
        array('created_at' => '2019-02-02', 'name' => 'Mary', 'id' => 2, 'age' => 20,),
        array('id' => 3, 'name' => '"Ethan', 'age' => 34),
        array('id' => 4, 'name' => '\'Tony', 'age' => 34),
        array('id' => 5, 'name' => ',', 'age' => 34),
        array('id' => 6, 'name' => '张三', 'age' => 34),
    );

    /**  \Closure $writer */
    $writer($data);
}, 'gbk');
