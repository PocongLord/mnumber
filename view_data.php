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
        // Setelah memasukkan data ke tab_purchasing
        $tabPurchasingStmt->close();

        // Siapkan prepared statement untuk tab_mrp
        $tabMrpStmt = $conn->prepare(
            "INSERT IGNORE INTO tab_mrp (
                mnumber_plant_sloc, mnumberplant, mrp1, mrp2, mrp3, storage, worksch1, worksch2, acc, cost, 
                mnumber, plant, sloc, comp_code, material_description, mrp_group, abc_indicator, 
                mrp_type, mrp_controller, lot_size, btci, msl, procurement_type, special_procurement_type, 
                indicator, issue_loc, default_external_pro_sloc, good_receipt_days, planned_delivery_days, 
                scheduling_margin_keys, safety_stock, strategy_group, consumption_mode, backward_period, 
                forward_period, avaibility_check, selection_method, tolerance_under_delivery, 
                unlimited_overdelivery, unit_of_issue, required_batch, sn_profile, valuation_class, 
                price_control, moving_average_price, standard_price, price_unit, do_not_cost, 
                quantity_structure, material_origin, overhead_group, variance_group, profit_center, 
                costing_lot_size
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        );

        if (!$tabMrpStmt) {
            die("Prepare failed for tab_mrp: " . $conn->error);
        }

        // Initialize $plantsData
        $plantsData = [];

        // Populate $plantsData with necessary data
        foreach ($plants as $plant) {
            $plantsData[] = [
                'mnumberplant' => $mnumber . $plant,
                'mnumber' => $mnumber,
                'plant' => $plant,
                'comp_code' => $comp_code,
                'material_description' => $material_description
            ];
        }

        // Loop untuk memasukkan data ke tab_mrp
foreach ($plantsData as $plantData) {
    $mnumberplant = $plantData['mnumberplant'];
    $mnumber = $plantData['mnumber'];
    $plant = $plantData['plant'];
    $comp_code = $plantData['comp_code'];
    $material_description = $plantData['material_description'];

    // Tentukan sloc berdasarkan plant
    $sloc = 'default_sloc'; // Replace with actual logic to determine sloc

    // Tentukan valuation_class berdasarkan material_type dari tab_create
    $materialTypeResult = $conn->query("SELECT material_type FROM tab_create WHERE mnumber = '$mnumber'");
    $materialTypeRow = $materialTypeResult->fetch_assoc();
    $material_type = $materialTypeRow['material_type'] ?? null;

    // Tentukan valuation_class
    $valuation_class = null;
    switch ($material_type) {
        case 'ZSPT':
            $valuation_class = 3000;
            break;
        case 'ZFUE':
            $valuation_class = 3010;
            break;
        case 'ZOIL':
            $valuation_class = 3020;
            break;
        case 'ZTIR':
            $valuation_class = 3030;
            break;
        case 'ZSFT':
            $valuation_class = 3040;
            break;
        case 'ZCON':
            $valuation_class = 3050;
            break;
    }

    // Tentukan profit_center
    $profit_center = ($plant === '0201') ? 'A2100' : 'A220A';

    // Concatenate mnumber, plant, and sloc into a variable
    $mnumber_plant_sloc = $mnumber . $plant . $sloc; // Store in a variable

    // Prepare variables for bind_param
    $mrp1 = 'X';
    $mrp2 = 'X';
    $mrp3 = 'X';
    $storage = 'X';
    $worksch1 = 'X';
    $worksch2 = 'X';
    $acc = 'X';
    $cost = 'X';
    $mrp_group = 'Z020';
    $mrp_type = 'PD';
    $lot_size = 'EX';
    $default_external_pro_sloc = '0001';
    $do_not_cost = 'X';
    $quantity_structure = 'X';
    $material_origin = 'X';
    
    // Define variables for all parameters that may be empty
    $abc_indicator = ''; // Set to empty string
    $btci = ''; // Set to empty string
    $msl = ''; // Set to empty string
    $special_procurement_type = ''; // Set to empty string
    $indicator = ''; // Set to empty string
    $issue_loc = ''; // Set to empty string
    $good_receipt_days = ''; // Set to empty string
    $planned_delivery_days = ''; // Set to empty string
    $safety_stock = ''; // Set to empty string
    $strategy_group = ''; // Set to empty string
    $consumption_mode = ''; // Set to empty string
    $backward_period = ''; // Set to empty string
    $forward_period = ''; // Set to empty string
    $selection_method = ''; // Set to empty string
    $tolerance_under_delivery = ''; // Set to empty string
    $unlimited_overdelivery = ''; // Set to empty string
    $unit_of_issue = ''; // Set to empty string
    $required_batch = ''; // Set to empty string
    $sn_profile = ''; // Set to empty string

    // Define scheduling_margin_keys as a variable
    $scheduling_margin_keys = '000'; // Set to '000' as a variable

    // Define procurement_type as a variable
    $procurement_type = 'F'; // Set to 'F' as a variable

    // Define availability_check as a variable
    $availability_check = '02'; // Set to '02' as a variable

    // Define price_control as a variable
    $price_control = 'V'; // Set to 'V' as a variable

    // Define moving_average_price, standard_price, and price_unit as empty strings
    $moving_average_price = ''; // Set to empty string
    $standard_price = ''; // Set to empty string
    $price_unit = ''; // Set to empty string

    // Define overhead_group as an empty string
    $overhead_group = ''; // Set to empty string

    // Define variance_group as a variable
    $variance_group = '000001'; // Set to '000001' as a variable

    // Define costing_lot_size as a variable
    $costing_lot_size = 1000; // Set to 1000 as a variable

    // Prepare the type definition and variables
    $type_definition = 'sssssssssssssssssssssssssssssssssssssssssssssssssss'; // 53 placeholders
    $variables = [
        $mnumber_plant_sloc, // mnumber_plant_sloc
        $mnumberplant, // mnumberplant
        $mrp1, // mrp1
        $mrp2, // mrp2
        $mrp3, // mrp3
        $storage, // storage
        $worksch1, // worksch1
        $worksch2, // worksch2
        $acc, // acc
        $cost, // cost
        $mnumber, // mnumber
        $plant, // plant
        $sloc, // sloc
        $comp_code, // comp_code
        $material_description, // material_description
        $mrp_group, // mrp_group
        $abc_indicator, // abc_indicator
        $mrp_type, // mrp_type
        $lot_size, // lot_size
        $btci, // btci
        $msl, // msl
        $procurement_type, // procurement_type
        $special_procurement_type, // special_procurement_type
        $indicator, // indicator
        $issue_loc, // issue_loc
        $default_external_pro_sloc, // default_external_pro_sloc
        $good_receipt_days, // good_receipt_days
        $planned_delivery_days, // planned_delivery_days
        $scheduling_margin_keys, // scheduling_margin_keys
        $safety_stock, // safety_stock
        $strategy_group, // strategy_group
        $consumption_mode, // consumption_mode
        $backward_period, // backward_period
        $forward_period, // forward_period
        $availability_check, // availability_check
        $selection_method, // selection_method
        $tolerance_under_delivery, // tolerance_under_delivery
        $unlimited_overdelivery, // unlimited_overdelivery
        $unit_of_issue, // unit_of_issue
        $required_batch, // required_batch
        $sn_profile, // sn_profile
        $valuation_class, // valuation_class
        $price_control, // price_control
        $moving_average_price, // moving_average_price
        $standard_price, // standard_price
        $price_unit, // price_unit
        $do_not_cost, // do_not_cost
        $quantity_structure, // quantity_structure
        $material_origin, // material_origin
        $overhead_group, // overhead_group
        $variance_group, // variance_group
        $profit_center, // profit_center
        $costing_lot_size // costing_lot_size
    ];

    // Debugging output
    echo "Type Definition Length: " . strlen($type_definition) . "<br>";
    echo "Number of Variables: " . count($variables) . "<br>";

    // Ensure the number of variables matches the type definition
    if (count($variables) !== strlen($type_definition)) {
        echo "Mismatch between type definition and number of variables.<br>";
        continue; // Skip this iteration if there's a mismatch
    }

    // Bind parameter untuk tab_mrp
    $tabMrpStmt->bind_param(
        $type_definition,
        ...$variables // Use the spread operator to pass the array as individual arguments
    );

    if (!$tabMrpStmt->execute()) {
        echo "Error inserting into tab_mrp: " . $tabMrpStmt->error . "<br>";
    }
}

      // Tutup prepared statement
        $tabMrpStmt->close();
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
