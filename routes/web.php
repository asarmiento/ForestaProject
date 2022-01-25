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
Route::get('/lista-de-fincas', [App\Http\Controllers\FarmController::class, 'listsFarms'])->name('listFarms');


Route::post('/data_route', [App\Http\Controllers\FarmController::class, 'importStore'])->name('data_route');
Route::post('/data_base', [App\Http\Controllers\DataBaseController::class, 'importStore'])->name('data_base');
Route::post('/data_familia', [App\Http\Controllers\DataBaseController::class, 'importFamily'])->name('data_familia');
Route::post('/reports-word', [App\Http\Controllers\DataBaseController::class, 'reportWord'])->name('reportWord');
Route::post('/reports-word-two', [App\Http\Controllers\DataBaseController::class, 'reportWordTwo'])->name('reportWordTwo');

