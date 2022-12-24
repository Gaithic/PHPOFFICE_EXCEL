<?php

use App\Http\Controllers\DynamicColumn;
use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ExcelController;
use App\Http\Controllers\FinalCode;

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
    return view('welcome');
});

Route::get('/excel', [ExcelController::class, 'excel'])->name('excel');
Route::get('/d_excel', [DynamicColumn::class, 'excel'])->name('d_excel');
Route::get('/f_excel', [FinalCode::class, 'excel'])->name('f_excel');