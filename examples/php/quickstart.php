<?php
require_once __DIR__ . '/vendor/autoload.php';

use OsbOffice\Osbxl\App;

$filePath = __DIR__ . '/invoice.xlsx';
if (file_exists($filePath)) {
    unlink($filePath);
}

$app = new App();
$wb = $app->createWorkBook($filePath);

try {
    $ws = $wb->addWorkSheet("Invoice");

    // Header
    $ws->getRange("A1")->setValue("Description");
    $ws->getRange("B1")->setValue("Amount");

    // Items
    $ws->getRange("A2")->setValue("System Implementation Consulting");
    $ws->getRange("B2")->setNumberValue(500000);

    $ws->getRange("A3")->setValue("Server Infrastructure Setup");
    $ws->getRange("B3")->setNumberValue(250000);

    $ws->getRange("A4")->setValue("Total");
    $ws->getRange("B4")->setFormula("SUM(B2:B3)");

    $wb->save();
    echo "Created: {$filePath}\n";
} finally {
    $wb->close();
}
