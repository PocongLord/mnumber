<?php
require 'vendor/autoload.php'; // Pastikan library PhpSpreadsheet terpasang

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Atur zona waktu ke waktu lokal pengupload
date_default_timezone_set('Asia/Jakarta'); // Ubah sesuai dengan zona waktu yang diinginkan

// Fungsi untuk membuat Material Number
function generateMaterialNumber($cellC, $cellE)
{
    if (stripos($cellC, 'Atlas Copco') !== false) {
        $prefixC = 'ACP';
    } elseif (stripos($cellC, 'Multi Flow') !== false || stripos($cellC, 'Multiflow') !== false) {
        $prefixC = 'MLF';
    } elseif (stripos($cellC, 'Manitau') !== false) {
        $prefixC = 'MAT';
    } else {
        $prefixC = substr(strtoupper($cellC), 0, 3);
    }

    $cleanedE = preg_replace('/[^A-Za-z0-9]/', '', strtoupper($cellE));
    $materialNumber = 'LG2' . $prefixC . $cleanedE;

    if (strlen($materialNumber) > 18) {
        $excessLength = strlen($materialNumber) - 18;
        $materialNumber = 'LG2' . $prefixC . substr($cleanedE, $excessLength);
    } elseif (strlen($materialNumber) < 18) {
        $remainingLength = 18 - strlen($materialNumber);
        $materialNumber = 'LG2' . $prefixC . str_pad($cleanedE, strlen($cleanedE) + $remainingLength, '0', STR_PAD_LEFT);
    }

    return strtoupper(substr($materialNumber, 0, 18));
}

// Proses Upload
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['file'])) {
    $file = $_FILES['file']['tmp_name'];

    if (empty($file)) {
        die("File is required.");
    }

    $spreadsheet = IOFactory::load($file);
    $sheet = $spreadsheet->getActiveSheet();

    // Iterasi melalui semua baris (mulai dari baris kedua)
    foreach ($sheet->getRowIterator(2) as $row) {
        $rowIndex = $row->getRowIndex();
        $cellC = strtoupper($sheet->getCell('C' . $rowIndex)->getValue() ?? '');
        $cellE = strtoupper($sheet->getCell('E' . $rowIndex)->getValue() ?? '');
        if (empty($cellC) || empty($cellE)) {
            continue;
        }

        // Generate material number
        $materialNumber = generateMaterialNumber($cellC, $cellE);
        $sheet->setCellValue('D' . $rowIndex, strtoupper($materialNumber));

        // Ubah seluruh kolom di baris menjadi huruf kapital
        foreach (range('A', $sheet->getHighestColumn()) as $col) {
            $value = $sheet->getCell($col . $rowIndex)->getValue();
            $sheet->setCellValue($col . $rowIndex, strtoupper($value ?? ''));
        }
    }

    // Simpan file hasil proses
    $outputDir = 'processed_files';
    if (!is_dir($outputDir)) {
        mkdir($outputDir, 0777, true);
    }

    // Ubah format nama file menjadi ddmmyyyyhhmm
    $newFileName = $outputDir . '/template-' . date('dmyHi') . '.xlsx';
    $writer = new Xlsx($spreadsheet);
    $writer->save($newFileName);

    // Kirim URL file hasil ke frontend
    echo json_encode(['file_url' => $newFileName]);
    exit;
}
?>
