<?php

include '../vendor/autoload.php';

use PFinal\Excel\Excel;

date_default_timezone_set('PRC');

$data = array(
    array('created' => '2015-01-01', 'product_id' => 873, 'quantity' => 1, 'amount' => '44.00', 'description' => 'misc'),
    array('created' => '2015-01-12', 'product_id' => 324, 'quantity' => 2, 'amount' => '88.00', 'description' => 'none'),
);


PFinal\Excel\Excel::exportExcel($data);