<?php

// Connect to the database
$database_name = 'example_database.xlsx'; // replace with your Excel database name
$worksheet_name = 'Sheet1'; // replace with the name of the worksheet where you want to insert the data
if (!file_exists($database_name)) {
    die("Error: The database file does not exist.");
}
$database = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
$database = \PhpOffice\PhpSpreadsheet\IOFactory::load($database_name);
$worksheet = $database->getSheetByName($worksheet_name);

// Get the form data
$name = $_POST['name'];
$year = $_POST['year'];
$place = $_POST['place'];

// Insert the data into the database
$row = $worksheet->getHighestRow() + 1; // get the next available row
$worksheet->setCellValue('A'.$row, $name);
$worksheet->setCellValue('B'.$row, $year);
$worksheet->setCellValue('C'.$row, $place);
$database_writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($database, 'Xlsx');
$database_writer->save($database_name);

// Display a success message
echo "Data inserted successfully.";

?>
