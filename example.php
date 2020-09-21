<?php
error_reporting(E_ALL);

require_once __DIR__ . '/vendor/autoload.php';
require(__DIR__."/big_arrray.php");

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

$spreadsheet = new Spreadsheet();
// $sheet = $spreadsheet->getActiveSheet();
// $sheet->setCellValue('A1', 'Hello World !');

// $writer = new Xlsx($spreadsheet);
// $writer->save('hello world.xlsx');
echo "<pre>";



$cab01 = array('uno','dos','tres');
$datos01 = array(1,2,3);
$datos02 = array([1,2,3]);
$datos03 = array([1,2,3],[11,22,33],[111,222,333]);
$datos04 = array(
    ['uno'=>1,'dos'=>2,'tres'=>3],
    ['roto'=>11,'roto'=>22,'roto'=>33],
    ['roto'=>111,'roto'=>222,'treezxceeeeeeeee'=>333],
    ['roto'=>11111,'dosxzczxsss'=>22222,'treeeeezxceeeeee'=>33333],
);
$datos05 = array(
    'uno00000'=>1,'dossss'=>2,'treeeeeeeeeee'=>3
);

// print_r(($datos01));
// print_r(($datos04));
// print_r(array_keys($datos04));
// print_r(($cab01));
$opt = [
    'filename'=> 'excel',
    'order_from_header' => true,
];
$excel = new Este($spreadsheet, $opt);

// dd($excel->headerStyle);
// return;

$multiAsoc = array( // ideal
    'pag 01' => [
        'uno','dos','tres'
    ],
    'pag 02' => [
        '1a1','2a2','3a3'
    ],
    'pag 03' => [
        't_tres','t_cuatro','t_por el'
    ]
);
$multiAsoc_ROTO = array( // ideal
    'pag XX1' => [
        'AAA', 'BBB', 'CCC'
    ],
    'pag XX2' => [
        'XXX', 'XXX', 'XXX'
    ],
    'NO TA' => [
        't_3333', 't_3333', 't_3333'
    ]
);
$multi = array(
    [
        'uno','dos','tres'
    ],
    [
        'aa','bb','cc'
    ],
    [
        '----','*****','////'
    ]
);
$headers03 = array( // ideal
    'pag 01' => [
        'uno','dos','tres'
    ]
);
$simple = array(
        'uno','dos','tres'
);
$datos01 = array(1,2,3,4,5,6);
$datos02 = array(['a','b','c'],[11,22,33],[111,222,333]);
$datos03 = array(
    'pag 01' =>
        ['uno'=>1,'dos'=>2,'tres'=>3],
    'pag 02' =>
        [['11' => 11, '22' => 22, '33' => 33], ['11' => 11, '22' => 22, '33' => 33], ['11' => 11, '22' => 22, '33' => 33]],
    'pag 03' =>
        ['tres'=>111,'cuatro'=>222,'por el'=>333],
    'NO TA' =>
        ['perro' => 111000, 'gato' => 222000, 'kaballo' => 3330000],
);
$datos03_b = array(
    'pag 01' =>
        ['uno' => 1, 'dos' => 2, 'tres' => 3],
    'pag 02' =>
        [['11' => 11, '22' => 22, '33' => 33], ['11' => 11, '22' => 22, '33' => 33], ['11' => 11, '22' => 22, '33' => 33]],
    'pag 03' =>
        ['tres' => 111, 'cuatro' => 222, 'por el' => 333],
    'NO TA' =>
        [111000, 222000, 3330000],
);
$datos04 = array(
    'uno00000'=>1,'dossss'=>2,'treeeeeeeeeee'=>3
);
$datos04_b = array(
    'pag 01' =>
        ['uno00000' => 1, 'dossss' => 2, 'treeeeeeeeeee' => 3]
);
$datos04_c = array(
    'pag 01' => [ 1,  2,  3]
);
// $excel->setHeader($headers);

$excel->generate($datos01,$simple);
// $excel->generate($datos02,$simple);
// $excel->generate($datos03,$simple);
// $excel->generate($datos03_b,$simple);
// $excel->generate($datos04,$simple);

// $excel->generate($datos01,$multiAsoc);
// $excel->generate($datos03,$multiAsoc);
// $excel->generate($datos03_b,$multiAsoc);
// $excel->generate($datos02,$multiAsoc);
// $excel->generate($datos04,$multiAsoc);

// $excel->generate($datos01,$multi);
// $excel->generate($datos03,$multi);
// $excel->generate($datos03_b,$multi);
// $excel->generate($datos02,$multi);
// $excel->generate($datos04,$multi);

// $excel->generate($datos01,$multiAsoc_ROTO);
// $excel->generate($datos03,$multiAsoc_ROTO);
// $excel->generate($datos03_b,$multiAsoc_ROTO);
// $excel->generate($datos02,$multiAsoc_ROTO);
// $excel->generate($datos04,$multiAsoc_ROTO);
// $excel->generate($datos04_b,$multiAsoc_ROTO);
// $excel->generate($datos04_c,$multiAsoc_ROTO);


















/*
$cab01 = array('uno','dos','tres');
$cab01_b = array('tres','dos','uno');
$cab02 = array('uno'=>'X','dos'=>'XX','tres'=>'XXX');
$datos01 = array(1,2,3);
$datos02 = array([1,2,3]);
$datos03 = array([1,2,3],[11,22,33],[111,222,333]);
$datos04 = array(
    ['uno'=>1,'dos'=>2,'tres'=>3],
    ['roto'=>11,'roto'=>22,'roto'=>33],
    ['roto'=>111,'roto'=>222,'roto'=>333],
    ['roto'=>11111,'dos'=>22222,'roto'=>33333],
);
$datos05 = array(
    'uno___'=>1,'dos___'=>2,'tres___'=>3
);
$datos06 = array(
    ['uno' => 1, 'dos' => 2, 'tres' => 3],
    ['uno' => 11, 'dos' => 22, 'tres' => 33],
    ['uno' => 111, 'dos' => 222, 'tres' => 333],
    ['uno' => 1111, 'dos' => 2222, 'tres' => 3333]
);
$excel->remove();
$excel->addHoja('titulo01'  , $datos01); // muestra 1,2,3
$excel->addHoja('titulo01_b', $datos02); // 1,2,3
$excel->addHoja('titulo02'  , $datos02, $cab01); //cab: un, dos, tres, sin datos cocincidentes
$excel->addHoja('titulo03'  , $datos03); // matriz sin cabecera
$excel->addHoja('titulo03_b', $datos03, $cab01); // solo la cabecera, sin datos coinc
$excel->addHoja('titulo04'  , $datos04, $cab01); // cab y el row que tiene la cab coinc
$excel->addHoja('titulo04_b', $datos04); // cab es el primer row, el body solo muestra los coinc
$excel->addHoja('titulo05'  , $datos05); // cabecera son las keys del row y el dato sus values
$excel->addHoja('titulo05_b', $datos05, $cab01); //cabecera indicada ningun dato coincidente
$excel->addHoja('titulo06'  , $datos06, $cab01); //cabecera indicada todo coincidente
$excel->addHoja('titulo06_b', $datos06, $cab01_b); // cabecera 2, body con orden segun cab
$excel->addHoja('titulo06_c', $datos06, $cab02); // cabecera es el value no las keys, nada coincidente
*/


$excel->objExcel->setActiveSheetIndex(0);
$writer = new Xlsx($excel->objExcel);
$writer->save('hello world.xlsx');

$excel->printLog();








function dd($val, $die=true) {
    echo '<pre>';
    switch (gettype($val)) {
        case 'string':
            if(strpos($val, strtoupper('SELECT')) !== false){
                echo SqlFormatter::format($val);
            } else {
                json_decode($val);
                if (json_last_error() == JSON_ERROR_NONE) {
                    echo "json";
                }
                echo $val;
            }
        break;
        case 'array':
        case 'object':
        default:
            array_map(function($val) {var_dump($val);}, func_get_args());
        break;
    }
    echo '</pre>';
    if($die) die();
}
