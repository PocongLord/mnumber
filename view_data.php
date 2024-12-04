<?php
require 'vendor/autoload.php'; // Pastikan library PhpSpreadsheet terpasang

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (!isset($_GET['file'])) {
    die('File not specified.');
}

$filePath = $_GET['file'];

// Validasi file
if (!file_exists($filePath)) {
    die('File not found.');
}

// Membaca file Excel
$spreadsheet = IOFactory::load($filePath);
$sheet = $spreadsheet->getActiveSheet();
$data = $sheet->toArray(null, true, true, true);
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>View Data</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {
            background-color: #f8f9fa;
        }
        .table-container {
            margin: 50px auto;
            max-width: 80%;
            overflow-x: auto;
        }
        table {
            white-space: nowrap;
        }
    </style>
</head>
<body>
    <div class="table-container">
        <h1 class="text-center">Data Excel</h1>
        <table class="table table-bordered table-striped">
            <thead class="table-dark">
                <tr>
                    <?php foreach ($data[1] as $header): ?>
                        <th><?php echo htmlspecialchars($header ?? '', ENT_QUOTES, 'UTF-8'); ?></th>
                    <?php endforeach; ?>
                </tr>
            </thead>
            <tbody>
                <?php foreach (array_slice($data, 1) as $row): ?>
                    <?php 
                    // Periksa apakah baris kosong
                    $isEmptyRow = true;
                    foreach ($row as $cell) {
                        if (!empty(trim($cell ?? ''))) {
                            $isEmptyRow = false;
                            break;
                        }
                    }                   
                    if ($isEmptyRow) {
                        continue; // Lewati baris kosong
                    }
                    ?>
                    <tr>
                        <?php foreach ($row as $cell): ?>
                            <td><?php echo htmlspecialchars($cell ?? '', ENT_QUOTES, 'UTF-8'); ?></td>
                        <?php endforeach; ?>
                    </tr>
                <?php endforeach; ?>
            </tbody>
        </table>
        
        <!-- Download Button -->
        <div class="d-flex justify-content-center gap-3">
    <a href="index.html" class="btn btn-primary">Kembali ke Menu Utama</a>
    <a href="processed_files/<?php echo basename($filePath); ?>" class="btn btn-success" download>Download Processed File</a>
</div>

    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
