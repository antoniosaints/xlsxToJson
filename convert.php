<?php

require 'vendor/autoload.php'; // Carrega o autoload do Composer

use PhpOffice\PhpSpreadsheet\IOFactory;

function excelToJson($excelFile)
{
    $spreadsheet = IOFactory::load($excelFile);
    $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

    $jsonArray = [];

    // Use a primeira linha como chaves
    $keys = array_shift($sheetData);
    
    // Itere sobre as linhas restantes
    foreach ($sheetData as $row) {
        // Verifique se a linha não está vazia
        if (!empty(array_filter($row))) {
            $rowData = [];

            // Associe os valores às chaves
            foreach ($row as $key => $value) {
                $rowData[$keys[$key]] = $value;
            }

            $jsonArray[] = $rowData;
        }
    }

    return json_encode($jsonArray, JSON_PRETTY_PRINT);
}

// Exemplo de uso
$excelFile = './contacts.xlsx';
$jsonOutput = excelToJson($excelFile);

// Salvar o JSON em um arquivo (opcional)
file_put_contents('output.json', $jsonOutput);

echo 'Conversão concluída.';
