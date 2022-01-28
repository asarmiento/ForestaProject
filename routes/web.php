<?php

use Illuminate\Support\Facades\Route;

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
    return view('auth/login');
});

Auth::routes();

Route::get('/home', [App\Http\Controllers\HomeController::class, 'index'])->name('home');
Route::get('/lista-de-familias', [App\Http\Controllers\FarmController::class, 'listsFamily'])->name('listFamily');
Route::get('/registro-de-familias', [App\Http\Controllers\FarmController::class, 'createFamily'])->name('get-family');
/**
 * Name Scientific
 */
Route::get('/lista-de-nombre-cientifico', [App\Http\Controllers\FarmController::class, 'listsScientific'])->name('listScientific');
Route::get('/registro-de-nombre-cientifico', [App\Http\Controllers\FarmController::class, 'createScientific'])->name('get-scientific');
Route::get('/registro-de-nombre-cientifico/{id}/edit', [App\Http\Controllers\FarmController::class, 'editScientific'])->name('edit-scientific');
Route::post('/store-scientific', [App\Http\Controllers\FarmController::class, 'storeScientific'])->name('store-scientific');
Route::put('/update-scientific/{id}', [App\Http\Controllers\FarmController::class, 'updateScientific'])->name('update-scientific');
/**
 * Name Common
 */
Route::get('/lista-de-nombre-comun', [App\Http\Controllers\FarmController::class, 'listsCommon'])->name('listCommon');
Route::get('/registro-de-nombre-comun', [App\Http\Controllers\FarmController::class, 'createCommon'])->name('get-common');
Route::get('/registro-de-nombre-comun/{id}/edit', [App\Http\Controllers\FarmController::class, 'editCommon'])->name('edit-common');
Route::post('/store-common', [App\Http\Controllers\FarmController::class, 'storeCommon'])->name('store-common');
Route::put('/update-common/{id}', [App\Http\Controllers\FarmController::class, 'updateCommon'])->name('update-common');
/**
 *
 */

Route::post('/store-family', [App\Http\Controllers\FarmController::class, 'storeFamily'])->name('store-family');
Route::get('/registro-de-nueva-finca', [App\Http\Controllers\FarmController::class, 'createFarm'])->name('get-farm');
Route::get('/registro-de-nueva-finca/{id}/edit', [App\Http\Controllers\FarmController::class, 'editFarm'])->name('edit-farm');
Route::put('/registro-de-nueva-finca/{id}', [App\Http\Controllers\FarmController::class, 'updateFarm'])->name('update-farm');
Route::get('/lista-de-fincas', [App\Http\Controllers\FarmController::class, 'listsFarms'])->name('listFarms');


Route::post('/data_route', [App\Http\Controllers\FarmController::class, 'importStore'])->name('data_route');
Route::post('/data_base', [App\Http\Controllers\DataBaseController::class, 'importStore'])->name('data_base');
Route::post('/data_familia', [App\Http\Controllers\DataBaseController::class, 'importFamily'])->name('data_familia');
Route::post('/reports-word', [App\Http\Controllers\DataBaseController::class, 'reportWord'])->name('reportWord');
Route::post('/reports-word-two', [App\Http\Controllers\DataBaseController::class, 'reportWordTwo'])->name('reportWordTwo');

