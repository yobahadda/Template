<?php
// Enable error logging
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

// Include PhpSpreadsheet library files
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

try {
    // Check if form is submitted
    if ($_SERVER["REQUEST_METHOD"] == "POST") {
        // Collect form data
        $name = isset($_POST['name']) ? $_POST['name'] : '';
        $email = isset($_POST['email']) ? $_POST['email'] : '';
        $subject = isset($_POST['subject']) ? $_POST['subject'] : '';
        $message = isset($_POST['message']) ? $_POST['message'] : '';

        // Load existing spreadsheet or create a new one
        $filePath = 'submissions.xlsx';
        if (file_exists($filePath)) {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($filePath);
        } else {
            $spreadsheet = new Spreadsheet();
            $spreadsheet->getActiveSheet()->setTitle('Submissions');
            // Set header row
            $spreadsheet->getActiveSheet()
                ->setCellValue('A1', 'Name')
                ->setCellValue('B1', 'Email')
                ->setCellValue('C1', 'Subject')
                ->setCellValue('D1', 'Message')
                ->setCellValue('E1', 'Submitted At');
        }

        // Get the active sheet
        $sheet = $spreadsheet->getActiveSheet();

        // Find the next empty row
        $row = $sheet->getHighestRow() + 1;

        // Write form data to the spreadsheet
        $sheet->setCellValue("A$row", $name)
              ->setCellValue("B$row", $email)
              ->setCellValue("C$row", $subject)
              ->setCellValue("D$row", $message)
              ->setCellValue("E$row", date('Y-m-d H:i:s'));

        // Save the spreadsheet
        $writer = new Xlsx($spreadsheet);
        $writer->save($filePath);

        echo 'Your message has been sent and saved. Thank you!';
    }
} catch (Exception $e) {
    echo 'Error: ' . $e->getMessage();
}
?>
