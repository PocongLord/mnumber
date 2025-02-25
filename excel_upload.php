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
    // Hapus spasi dari cellC  
    $cellC = str_replace(' ', '', $cellC);  
  
    // Tentukan prefix berdasarkan nilai di cellC  
    if (stripos($cellC, 'AtlasCopco') !== false) {  
        $prefixC = 'ACP';  
    } elseif (stripos($cellC, 'MultiFlow') !== false) {  
        $prefixC = 'MLF';  
    } elseif (stripos($cellC, 'Manitau') !== false) {  
        $prefixC = 'MAT';  
    } elseif (stripos($cellC, 'Manitou') !== false) {  
        $prefixC = 'MAT';  
    } else {  
        $prefixC = substr(strtoupper($cellC), 0, 3);  
    }  
  
    // Bersihkan cellE dari karakter non-alphanumeric dan konversi ke uppercase  
    $cleanedE = preg_replace('/[^A-Za-z0-9]/', '', strtoupper($cellE));  
  
    // Tambahkan leading zeros untuk memastikan panjang total 18 karakter  
    $materialNumber = 'LG2' . $prefixC . str_pad($cleanedE, 12, '0', STR_PAD_LEFT);  
  
    // Pastikan hasil akhir adalah 18 karakter  
    return strtoupper(substr($materialNumber, 0, 18));  
}  
  
// Proses Upload  
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['file'])) {  
    $file = $_FILES['file']['tmp_name'];  
  
    if (empty($file)) {  
        die(json_encode(['error' => 'File is required.']));  
    }  
  
    try {  
        $spreadsheet = IOFactory::load($file);  
        $sheet = $spreadsheet->getActiveSheet();  
  
        // Ubah header kolom D menjadi mnumber  
        $headerRowIndex = 1; // Baris header, biasanya baris pertama  
        foreach ($sheet->getRowIterator($headerRowIndex, $headerRowIndex) as $headerRow) {  
            $cellD = $sheet->getCell('D' . $headerRow->getRowIndex());  
            if (strtoupper($cellD->getValue() ?? '') === '3 DIGIT MANUFAKTUR') {  
                $sheet->setCellValue('D' . $headerRow->getRowIndex(), 'mnumber');  
            }  
        }  
  
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
    } catch (\Exception $e) {  
        die(json_encode(['error' => 'An error occurred: ' . $e->getMessage()]));  
    }  
} else {  
    die(json_encode(['error' => 'Invalid request method or file not provided.']));  
}  
?>  
