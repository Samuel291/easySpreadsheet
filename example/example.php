<?php

    require '../vendor/autoload.php';

    use Samuel291\EasySpreadsheet;

    $spreadsheet = new EasySpreadsheet();

    //To insert a new tab in the worksheet, copy and paste the code snippet below before the save
    $spreadsheet->startSheet('NomeJanela01');
    $spreadsheet->setHeader('Coluna 01', ['cellColor' => 'FFFF00']);
    $spreadsheet->setItem('Conteudo 01', ['fontColor' => '000FF0']);
    $spreadsheet->setContent();
    $spreadsheet->render();

    $spreadsheet->save("NomedoArquivo");