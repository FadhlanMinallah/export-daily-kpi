<?php

namespace App\Http\Controllers;

use App\Exports\DailyKPIExport;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;

class ReportController extends Controller
{
    public function downloadExcel(Request $request)
    {
        $tanggal = $request->input('tanggal', date('Y-m-d'));
        $lokasi = $request->input('lokasi', 'ID_SITE');
        
        // Ambil data dari database atau hardcode untuk testing
        $data = [
            'gps_terpasang' => 30,
            'gps_online' => 20,
            'gps_offline' => 10,
            'gps_online_persen' => 67,
            'dashcam_terpasang' => 30,
            'dashcam_online' => 16,
            'dashcam_offline' => 14,
            'dashcam_online_persen' => 53,
        ];
        
        $fileName = 'Report_Daily_KPI' . date('Ymd', strtotime($tanggal)) . '.xlsx';
        
        return Excel::download(
            new DailyKPIExport($tanggal, $lokasi, $data),
            $fileName,
            \Maatwebsite\Excel\Excel::XLSX,
            [
                'includeCharts' => true,
            ]
        );
    }
}