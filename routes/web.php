<?php

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

Route::get('/daily-kpi-export', 'ReportController@downloadExcel')->name('report.download.excel');

Route::get('/timely-kpi-export', 'ReportController@downloadTimelyReporting')->name('report.download.excel.timely');

Route::get('/summary-export', 'ReportController@downloadSummaryExport')->name('report.download.summary');

Route::get('/nilai-kpi-export', 'ReportController@downloadNilaiKpi')->name('report.download.nilai-kpi');

Route::get('/timely-reporting-sum', 'ReportController@downloadTimelyReportingSum')->name('report.download.timely-sum');

Route::get('/bapp-export', 'ReportController@downloadBappExport')->name('report.download.bapp');

Route::get('/kegagalan-penormalan-export', 'ReportController@downloadKegagalanPenormalan')->name('report.download.kegagalan-penormalan');

Route::get('/kpi-export', 'ReportController@downloadKpiExport')->name('report.download.kpi');

Route::get('/response-time-export', 'ReportController@downloadResponseTimeExport')->name('report.download.response-time');

Route::get('/lampiran-nopol-export', 'ReportController@downloadLampiranNopolExport')->name('report.download.lampiran-nopol');