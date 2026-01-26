<?php

namespace App\Http\Controllers;

use App\Exports\DailyKPIExport;
use App\Exports\SummaryExport;
use App\Exports\TimelyReportingExport;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;

class ReportController extends Controller
{
    public function downloadExcel(Request $request)
    {
        // Ambil data dari database atau hardcode untuk testing
        $data = [
            'date'                  => '2026-01-23',
            'site'                  => 'MT Surabaya',
            'gps_terpasang'         => 30,
            'gps_online'            => 20,
            'gps_offline'           => 10,
            'gps_online_persen'     => 67,
            'dashcam_terpasang'     => 30,
            'dashcam_online'        => 16,
            'dashcam_offline'       => 14,
            'dashcam_online_persen' => 53,
        ];
        
        $fileName = 'Report_Daily_KPI' . date('Ymd', strtotime($data['date'])) . '.xlsx';
        
        return Excel::download(
            new DailyKPIExport($data, 'daily_kpi'),
            $fileName,
            \Maatwebsite\Excel\Excel::XLSX,
            [
                'includeCharts' => true,
            ]
        );
    }
    
    public function downloadTimelyReporting()
    {
        $data = [
            'periode' => 'Jan-26',
            'site' => 'IT SURABAYA GROUP',

            'items' => [

                /* ===================== STATUS PERANGKAT ===================== */
                [
                    'item' => 'Perangkat Status Offline GPS',
                    'keterangan' => 'Persentase Perangkat Status Offline GPS',
                    'is_percent' => true,
                    'values' => [
                        1=>0.05,2=>0.04,3=>0.04,4=>0.05,5=>0.04,6=>0.04,7=>0.04,
                        8=>0.04,9=>0.03,10=>0.04,11=>0.04,12=>0.03,13=>0.03,
                        14=>0.04,15=>0.03,16=>0.04,17=>0.04,18=>0.04,19=>0.04,
                        20=>0.00
                    ]
                ],
                [
                    'item' => 'Perangkat Status Offline Dashcam',
                    'keterangan' => 'Persentase Perangkat Status Offline Dashcam',
                    'is_percent' => true,
                    'values' => [
                        1=>0.09,2=>0.08,3=>0.07,4=>0.08,5=>0.08,6=>0.06,7=>0.06,
                        8=>0.07,9=>0.05,10=>0.06,11=>0.06,12=>0.07,13=>0.07,
                        14=>0.05,15=>0.04,16=>0.06,17=>0.05,18=>0.08,19=>0.06,
                        20=>0.00
                    ]
                ],
                [
                    'item' => 'Perangkat Status Online',
                    'keterangan' => 'Persentase Perangkat Status Online',
                    'is_percent' => true,
                    'values' => [
                        1=>0.95,2=>0.96,3=>0.96,4=>0.95,5=>0.96,6=>0.96,7=>0.96,
                        8=>0.96,9=>0.97,10=>0.96,11=>0.96,12=>0.97,13=>0.97,
                        14=>0.96,15=>0.97,16=>0.96,17=>0.96,18=>0.96,19=>0.96,
                        20=>1.00
                    ]
                ],

                /* ===================== ARMADA ===================== */
                [
                    'item' => 'Jumlah Mobil Tangki dengan Perangkat Terinstal',
                    'keterangan' => 'Jumlah Perangkat Terinstal shipment harian',
                    'is_percent' => false,
                    'values' => [
                        1=>225,2=>225,3=>225,4=>226,5=>226,6=>226,7=>226,8=>225,
                        9=>225,10=>225,11=>225,12=>225,13=>225,14=>225,15=>225,
                        16=>225,17=>225,18=>225,19=>225,20=>9
                    ]
                ],
                [
                    'item' => 'Jumlah Mobil Tangki dengan Perangkat Offline GPS',
                    'keterangan' => 'Jumlah Perangkat Offline GPS dalam shipment harian',
                    'is_percent' => false,
                    'values' => [
                        1=>12,2=>10,3=>8,4=>11,5=>9,6=>9,7=>10,8=>8,9=>6,
                        10=>8,11=>8,12=>7,13=>7,14=>9,15=>7,16=>9,17=>8,
                        18=>10,19=>9,20=>0
                    ]
                ],
                [
                    'item' => 'Jumlah Mobil Tangki dengan Perangkat Offline',
                    'keterangan' => 'Jumlah Perangkat Offline Dashcam shipment harian',
                    'is_percent' => false,
                    'values' => [
                        1=>20,2=>18,3=>15,4=>18,5=>17,6=>14,7=>13,8=>12,
                        9=>13,10=>13,11=>14,12=>15,13=>12,14=>10,15=>13,
                        16=>13,17=>10,18=>17,19=>13,20=>0
                    ]
                ],

                /* ===================== SAFETY ===================== */
                [
                    'item' => 'Jumlah Kegagalan Integrasi SIDO',
                    'keterangan' => 'Jumlah kegagalan integrasi SIDO',
                    'is_percent' => false,
                    'values' => [
                        1=>0,2=>0,3=>0,4=>0,5=>0,6=>0,7=>0,8=>0,9=>0,
                        10=>0,11=>0,12=>0,13=>0,14=>0,15=>0,16=>0,
                        17=>0,18=>0,19=>0,20=>0
                    ]
                ],
                [
                    'item' => 'Layanan 24 Jam',
                    'keterangan' => 'Cakupan layanan 24 jam GPS dan Dashcam',
                    'is_percent' => true,
                    'values' => [
                        1=>1,2=>1,3=>1,4=>1,5=>1,6=>1,7=>1,8=>1,9=>1,
                        10=>1,11=>1,12=>1,13=>1,14=>1,15=>1,16=>1,
                        17=>1,18=>1,19=>1,20=>1
                    ]
                ],
                [
                    'item' => 'Pelanggaran Driver',
                    'keterangan' => 'Jumlah pelanggaran dari ANFA dan TM',
                    'is_percent' => false,
                    'values' => [
                        1=>1947,2=>2015,3=>2306,4=>2654,5=>2340,6=>2445,
                        7=>2252,8=>2694,9=>3047,10=>3154,11=>3094,
                        12=>1933,13=>2837,14=>3190,15=>2683,16=>3779,
                        17=>3317,18=>2940,19=>2654,20=>0
                    ]
                ],

                /* ===================== KEJADIAN ===================== */
                [
                    'item' => 'Pelanggaran Driving Behavior dari DMS (Kejadian)',
                    'keterangan' => 'Total pelanggaran driver behavior',
                    'is_percent' => false,
                    'values' => [
                        1=>19164,2=>19785,3=>22780,4=>26136,5=>23017,
                        6=>23796,7=>22037,8=>26314,9=>29613,10=>30901,
                        11=>30300,12=>18884,13=>27684,14=>31162,
                        15=>28079,16=>36767,17=>32621,18=>28719,
                        19=>25463,20=>0
                    ]
                ],
                [
                    'item' => 'Harsh Braking (Kejadian)',
                    'keterangan' => 'Total pelanggaran pengereman mendadak',
                    'is_percent' => false,
                    'values' => [
                        1=>175,2=>194,3=>180,4=>234,5=>178,6=>301,
                        7=>205,8=>208,9=>307,10=>214,11=>325,
                        12=>184,13=>292,14=>271,15=>185,16=>404,
                        17=>225,18=>287,19=>302,20=>0
                    ]
                ],
                [
                    'item' => 'Jumlah Harsh Acceleration (Kejadian)',
                    'keterangan' => 'Total pelanggaran peningkatan kecepatan mendadak',
                    'is_percent' => false,
                    'values' => [
                        1=>131,2=>140,3=>100,4=>143,5=>122,6=>265,
                        7=>109,8=>167,9=>293,10=>188,11=>244,
                        12=>252,13=>193,14=>312,15=>160,16=>190,
                        17=>326,18=>0
                    ]
                ],
                [
                    'item' => 'Jumlah Harsh Cornering (Kejadian)',
                    'keterangan' => 'Total belok mendadak',
                    'is_percent' => false,
                    'values' => [
                        1=>3,2=>31,3=>4,4=>3,5=>48,6=>70,7=>178,
                        8=>252,9=>235,10=>216,11=>10,12=>113,
                        13=>190,14=>265,15=>220,16=>218,17=>164,
                        18=>193,19=>267,20=>0
                    ]
                ],
                [
                    'item' => 'Jumlah Over Speed (Kejadian)',
                    'keterangan' => 'Total pelanggaran melebihi batas kecepatan 60 km/jam',
                    'is_percent' => false,
                    'values' => [
                        1=>1,2=>2,3=>4,4=>27,5=>28,6=>12,7=>6,
                        8=>7,9=>17,10=>8,11=>55,12=>8,13=>15,
                        14=>13,15=>6,16=>8,17=>26,18=>0
                    ]
                ],
            ]
        ];

        $fileName = 'Timely_Reporting_' . $data['periode'] . '.xlsx';
        
        return Excel::download(
            new TimelyReportingExport($data, 'timely_reporting'),
            $fileName,
            \Maatwebsite\Excel\Excel::XLSX,
        );
    }

    public function downloadSummaryExport()
    {
        $fileName = 'Summary_Export.xlsx';
        $data = [];
        
        return Excel::download(
            new SummaryExport($data, 'Summary'),
            $fileName,
            \Maatwebsite\Excel\Excel::XLSX,
        );
    }
}