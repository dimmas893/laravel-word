<?php

use Illuminate\Support\Facades\Route;
use Novay\WordTemplate\WordTemplate;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/', function () {
    $wordTest = new \PhpOffice\PhpWord\PhpWord();

    $newSection = $wordTest->addSection();

    $desc1 = "The Portfolio details is a very useful feature of the web page. You can establish your archived details and the works to the entire web community. It was outlined to bring in extra clients, get you selected based on this details.";

    $newSection->addText($desc1, array('name' => 'Tahoma', 'size' => 15));

    $objectWriter = \PhpOffice\PhpWord\IOFactory::createWriter($wordTest, 'Word2007');
    try {
        $objectWriter->save(storage_path('TestWordFile.docx'));
    } catch (Exception $e) {
    }
    return response()->download(storage_path('TestWordFile.docx'));
});
