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

// Fungsi untuk memasukkan data ke database
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['insert_to_db'])) {
    $host = 'localhost'; // Sesuaikan dengan host database Anda
    $username = 'root';  // Sesuaikan dengan username database Anda
    $password = '';      // Sesuaikan dengan password database Anda
    $dbname = 'db_mnumber'; // Sesuaikan dengan nama database Anda

    $conn = new mysqli($host, $username, $password, $dbname);

    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    }

    // Hapus semua data di tabel material dan tab_create sebelum memasukkan data baru
    $conn->query("TRUNCATE TABLE material");
    $conn->query("TRUNCATE TABLE tab_create");
    $conn->query("TRUNCATE TABLE tab_uom");
    $conn->query("TRUNCATE TABLE tab_purchasing");

    // Periksa apakah query berhasil dijalankan
    if ($conn->error) {
        die("Failed to clear tables: " . $conn->error);
    }

    $stmt = $conn->prepare(
        "INSERT INTO material (manufaktur, mnumber, old_material_number, material_description, material_group, external_material_group, material_type, uom, comp_code, keterangan, systodb) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW())"
    );

    if (!$stmt) {
        die("Prepare failed: " . $conn->error);
    }

    // Mulai proses loop untuk memasukkan data baru
    foreach (array_slice($data, 1) as $row) {
        if (empty(trim($row['A'] ?? ''))) continue; // Lewati baris kosong

        // Ambil data dari kolom yang sesuai
        $manufaktur = $row['C'] ?? '';
        $mnumber = $row['D'] ?? '';
        $old_material_number = $row['E'] ?? '';
        $material_description = $row['F'] ?? '';
        $material_group = $row['G'] ?? '';
        $external_material_group = $row['H'] ?? '';
        $material_type = $row['I'] ?? '';
        $uom = $row['J'] ?? '';
        $comp_code = $row['K'] ?? '';
        $keterangan = $row['L'] ?? ''; // Ambil data dari kolom L

        // Bind parameter untuk tabel material
        $stmt->bind_param(
            'ssssssssss',
            $manufaktur,
            $mnumber,
            $old_material_number,
            $material_description,
            $material_group,
            $external_material_group,
            $material_type,
            $uom,
            $comp_code,
            $keterangan
        );

        if (!$stmt->execute()) {
            echo "Error inserting row: " . $stmt->error . "<br>";
        } else {
            // Jika berhasil memasukkan ke tabel material, masukkan ke tabel tab_create
            $tabCreateStmt = $conn->prepare(
                "INSERT INTO tab_create (basic, mnumber, industry_sector, material_type, material_description, uom, material_group, old_material_number, gross_weight, net_weight, weight_unit)
                VALUES ('X', ?, 'Z', ?, ?, ?, ?, ?, NULL, NULL, 'KG')"
            );

            if (!$tabCreateStmt) {
                echo "Prepare failed for tab_create: " . $conn->error . "<br>";
                continue;
            }

            $tabCreateStmt->bind_param(
                'ssssss', // 6 parameter sesuai query
                $mnumber,
                $material_type,
                $material_description,
                $uom,
                $material_group,
                $old_material_number
            );

            if (!$tabCreateStmt->execute()) {
                echo "Error inserting into tab_create: " . $tabCreateStmt->error . "<br>";
            }

            $tabCreateStmt->close();
        }

       // Logika untuk insert ke tab_purchasing
// Tentukan plant berdasarkan COMP CODE
$plants = [];
if ($comp_code === '0200') {
    $plants = ['0201', '0202', '0203'];
} elseif ($comp_code === '0100') {
    $plants = ['0101', '0102', '0103'];
}

// Insert ke tabel tab_purchasing untuk setiap plant
$tabPurchasingStmt = $conn->prepare(
    "INSERT IGNORE INTO tab_purchasing (mnumberplant, purchasing, mnumber, plant, comp_code, material_description, purchasing_group, purchasing_variable_active) 
    VALUES (?, 'X', ?, ?, ?, ?, '000', '1')"
);

if (!$tabPurchasingStmt) {
    echo "Prepare failed for tab_purchasing: " . $conn->error . "<br>";
    exit;
}

// Proses loop untuk setiap plant
foreach ($plants as $plant) {
    // Gabungkan mnumber dan plant untuk mnumberplant
    $mnumberplant = $mnumber . $plant;

    // Lakukan bind_param dan execute untuk setiap plant
    $tabPurchasingStmt->bind_param(
        'sssss', // Format string untuk 5 parameter
        $mnumberplant,  // mnumberplant (gabungan mnumber dan plant)
        $mnumber,       // mnumber
        $plant,         // plant
        $comp_code,     // comp_code
        $material_description // material_description
    );

    if (!$tabPurchasingStmt->execute()) {
        echo "Error inserting into tab_purchasing: " . $tabPurchasingStmt->error . "<br>";
    }
}
$tabPurchasingStmt->close();


    }

    // Tambahkan logika untuk insert ke tabel tab_uom
    $tabUomStmt = $conn->prepare(
        "INSERT IGNORE INTO tab_uom (basic, mnumber, material_description, denominator, alternatif_uom, numerator, length, width, height, unit_dimension, volume, volume_unit) 
         VALUES ('X', ?, ?, '1', ?, '1', NULL, NULL, NULL, NULL, NULL, NULL)"
    );

    if (!$tabUomStmt) {
        die("Prepare failed for tab_uom: " . $conn->error);
    }

    // Proses loop untuk memasukkan data ke tab_uom
    foreach (array_slice($data, 1) as $row) {
        if (empty(trim($row['A'] ?? ''))) continue; // Lewati baris kosong

        // Ambil data yang dibutuhkan dari kolom
        $mnumber = $row['D'] ?? '';
        $material_description = $row['F'] ?? '';
        $alternatif_uom = $row['J'] ?? '';

        // Bind parameter untuk tab_uom
        $tabUomStmt->bind_param(
            'sss', // 3 parameter sesuai query
            $mnumber,
            $material_description,
            $alternatif_uom
        );

        if (!$tabUomStmt->execute()) {
            echo "Error inserting into tab_uom: " . $tabUomStmt->error . "<br>";
        }
    }

    // Tutup prepared statement
    $tabUomStmt->close();

    $stmt->close();
    $conn->close();
    echo "<div class='alert alert-success'>Data berhasil dimasukkan ke database.</div>";
}
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
        .already-exists {
            background-color: #ffcccc; /* Warna merah muda untuk menandakan sudah ada */
        }
    </style>
</head>
<body>
    <div class="table-container">
        <h1 class="text-center">Data Excel</h1>
        <form method="POST">
            <table class="table table-bordered table-striped">
                <thead class="table-dark">
                    <tr>
                        <?php foreach ($data[1] as $header): ?>
                            <th><?php echo htmlspecialchars($header ?? '', ENT_QUOTES, 'UTF-8'); ?></th>
                        <?php endforeach; ?>
                    </tr>
                </thead>
                <tbody>
                    <?php
                    // Tampilkan data dalam tabel
                    foreach (array_slice($data, 1) as $row): ?>
                        <tr>
                            <?php foreach ($row as $cell): ?>
                                <td><?php echo htmlspecialchars($cell ?? '', ENT_QUOTES, 'UTF-8'); ?></td>
                            <?php endforeach; ?>
                        </tr>
                    <?php endforeach; ?>
                </tbody>
            </table>
            <div class="d-flex justify-content-center gap-3">
                <a href="index.html" class="btn btn-primary">Kembali ke Menu Utama</a>
                <a href="processed_files/<?php echo basename($filePath); ?>" class="btn btn-success" download>Download Processed File</a>
                <button type="submit" name="insert_to_db" class="btn btn-warning">Input ke Database</button>
            </div>
        </form>
    </div>
</body>
</html>
