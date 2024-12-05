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

    $stmt = $conn->prepare(
        "INSERT INTO material (manufaktur, mnumber, old_material_number, material_description, material_group, external_material_group, material_type, uom, comp_code, keterangan, systodb) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW())
        ON DUPLICATE KEY UPDATE 
        manufaktur = VALUES(manufaktur),
        old_material_number = VALUES(old_material_number),
        material_description = VALUES(material_description),
        material_group = VALUES(material_group),
        external_material_group = VALUES(external_material_group),
        material_type = VALUES(material_type),
        uom = VALUES(uom),
        comp_code = VALUES(comp_code),
        keterangan = VALUES(keterangan)"
    );
    
    if (!$stmt) {
        die("Prepare failed: " . $conn->error);
    }
    
    $duplicateMnumbers = []; // Menyimpan mnumber yang sudah ada
    $duplicateMnumbersMessage = '';

    // Loop untuk memasukkan data
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

    // Cek apakah mnumber sudah ada di database
    $checkStmt = $conn->prepare("SELECT COUNT(*) FROM material WHERE mnumber = ?");
    $checkStmt->bind_param('s', $mnumber);
    $checkStmt->execute();
    $checkStmt->bind_result($count);
    $checkStmt->fetch();
    $checkStmt->close();

    if ($count > 0) {
        // Jika data sudah ada, simpan mnumber dan beri notifikasi
        $duplicateMnumbers[] = $mnumber;
        continue; // Lewati baris ini, tidak memasukkan ke database
    }

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
            VALUES ('x', ?, 'Z', ?, ?, ?, ?, ?, NULL, NULL, 'KG')"
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
}

    
    // Menampilkan notifikasi jika ada data duplikat
    if (count($duplicateMnumbers) > 0) {
        $duplicateMnumbersMessage = "Ada data yang sudah pernah input: " . implode(', ', $duplicateMnumbers);
    }
    
    $stmt->close();
    $conn->close();

    if ($duplicateMnumbersMessage) {
        echo "<div class='alert alert-warning'>{$duplicateMnumbersMessage}</div>";
    } else {
        echo "<div class='alert alert-success'>Data berhasil dimasukkan ke database.</div>";
    }
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
                    // Koneksi ke database untuk pengecekan duplikat mnumber
                    $conn = new mysqli('localhost', 'root', '', 'db_mnumber');
                    if ($conn->connect_error) {
                        die("Connection failed: " . $conn->connect_error);
                    }

                    // Ambil semua mnumber yang sudah ada di database
                    $existingMnumbers = [];
                    $result = $conn->query("SELECT mnumber FROM material");
                    if ($result) {
                        while ($row = $result->fetch_assoc()) {
                            $existingMnumbers[] = $row['mnumber'];
                        }
                    }

                    // Menampilkan data
                    foreach (array_slice($data, 1) as $row): 
                        $mnumber = $row['D'] ?? '';
                        $rowClass = in_array($mnumber, $existingMnumbers) ? 'already-exists' : '';
                    ?>
                        <tr class="<?php echo $rowClass; ?>">
                            <?php foreach ($row as $cell): ?>
                                <td><?php echo htmlspecialchars($cell ?? '', ENT_QUOTES, 'UTF-8'); ?></td>
                            <?php endforeach; ?>
                        </tr>
                    <?php endforeach; ?>

                    <?php $conn->close(); ?>
                </tbody>
            </table>
            <!-- Tombol aksi -->
            <div class="d-flex justify-content-center gap-3">
                <a href="index.html" class="btn btn-primary">Kembali ke Menu Utama</a>
                <a href="processed_files/<?php echo basename($filePath); ?>" class="btn btn-success" download>Download Processed File</a>
                <button type="submit" name="insert_to_db" class="btn btn-warning">Input ke Database</button>
            </div>
        </form>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
