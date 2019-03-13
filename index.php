<?php
/**
 * Created by PhpStorm.
 * User: lea
 * Date: 2017/11/2
 * Time: 14:37
 */

require './vendor/autoload.php';

$db_cfg = [
    // 数据库类型
    'type'     => 'mysql',
    // 服务器地址
    'hostname' => '192.168.33.10',
    // 数据库名
    'database' => 'mdw',
    // 用户名
    'username' => 'root',
    // 密码
    'password' => '123456',
    // 端口
    'hostport' => '3306',
];

try {
    $mysql = new PDO("{$db_cfg['type']}:host={$db_cfg['hostname']};dbname={$db_cfg['database']}", $db_cfg['username'], $db_cfg['password']);
    $mysql->query("SET NAMES utf8mb4");
} catch (PDOException $e) {
    exit('数据库连接错误！错误信息：' . $e->getMessage());
}
//查看表
$res    = $mysql->query('SHOW TABLE STATUS');
$tables = [];
while ($row = $res->fetch()) {
    array_push($tables, [
        'name'      => $row['Name'],
        'engine'    => $row['Engine'],
        'collation' => $row['Collation'],
        'comment'   => $row['Comment'],
    ]);
}

//查询表
foreach ($tables as &$val) {
    $res    = $mysql->query("SHOW FULL FIELDS FROM `{$val['name']}`");
    $fields = [];
    while ($row = $res->fetch()) {
        array_push($fields, [
            'Field'     => $row['Field'],
            'Type'      => $row['Type'],
            'Collation' => $row['Collation'],
            'Null'      => $row['Null'],
            'Key'       => $row['Key'],
            'Default'   => $row['Default'],
            'Extra'     => $row['Extra'],
            'Comment'   => $row['Comment'],
        ]);
    }
    $val['field'] = $fields;
}

$excel = new PHPExcel();
$excel->getProperties()->setCreator('lea<cotyxpp@qq.com>');
$excel->getProperties()->setTitle("数据字典信息");

$excel->getDefaultStyle()->getFont()->setName('宋体')->setSize(10);

$excel->setActiveSheetIndex(0);
$excel->getActiveSheet()->setTitle('数据字典');
$activeSheet = $excel->getActiveSheet();

$activeSheet->getColumnDimension('B')->setWidth(10);
$activeSheet->getColumnDimension('C')->setWidth(20);
$activeSheet->getColumnDimension('D')->setWidth(24);
$activeSheet->getColumnDimension('E')->setWidth(20);
$activeSheet->getColumnDimension('F')->setWidth(12);
$activeSheet->getColumnDimension('G')->setWidth(12);
$activeSheet->getColumnDimension('H')->setWidth(18);
$activeSheet->getColumnDimension('I')->setWidth(30);

$activeSheet->getDefaultStyle()->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$activeSheet->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$activeSheet->getDefaultRowDimension()->setRowHeight(20);
$styleArray = [
    'borders' => [
        'allborders' => [
            //'style' => PHPExcel_Style_Border::BORDER_THICK,//边框是粗的
            'style' => PHPExcel_Style_Border::BORDER_THIN, //细边框
            //'color' => ['argb' => 'FFFF0000'],
        ],
    ],
];

$num = 1;
foreach ($tables as $key => $val) {
    $activeSheet->setCellValue('A' . $num, '表' . ($key + 1) . ' ' . $val['name'] . ($val['comment'] ? ' （' . $val['comment'] . '）' : ''));
    $activeSheet->mergeCells('A' . $num . ':I' . $num);
    $num++;

    $start = $num;
    $activeSheet->setCellValue('B' . $num, '序号');
    $activeSheet->setCellValue('C' . $num, '字段名');
    $activeSheet->setCellValue('D' . $num, '类型（长度）');
    $activeSheet->setCellValue('E' . $num, '字符集');
    $activeSheet->setCellValue('F' . $num, 'Null');
    $activeSheet->setCellValue('G' . $num, 'Key');
    $activeSheet->setCellValue('H' . $num, '其它');
    $activeSheet->setCellValue('I' . $num, '备注');
    $num++;
    foreach ($val['field'] as $k => $v) {
        $activeSheet->setCellValue('B' . $num, $k + 1);
        $activeSheet->setCellValue('C' . $num, $v['Field']);
        $activeSheet->setCellValue('D' . $num, $v['Type']);
        $activeSheet->setCellValue('E' . $num, $v['Collation']);
        $activeSheet->setCellValue('F' . $num, $v['Null']);
        $activeSheet->setCellValue('G' . $num, $v['Key']);
        $activeSheet->setCellValue('H' . $num, $v['Extra']);
        $activeSheet->setCellValue('I' . $num, $v['Comment']);
        $num++;
    }
    $activeSheet->getStyle('B' . $start . ':I' . ($num - 1))->applyFromArray($styleArray);
    $num++;
}
$write = new PHPExcel_Writer_Excel2007($excel);
$write->save("数据字典".date('YmdHis').".xlsx");
