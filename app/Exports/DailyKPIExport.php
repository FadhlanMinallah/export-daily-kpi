<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithCharts;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;

use PhpOffice\PhpSpreadsheet\Chart\Layout;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

use PhpOffice\PhpSpreadsheet\Chart\Chart;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class DailyKPIExport implements FromCollection, WithEvents, WithColumnWidths, WithCharts, WithTitle
{
    protected string $tanggal;
    protected string $lokasi;
    protected array $data;

    public function __construct(string $tanggal, string $lokasi, array $data)
    {
        $this->tanggal = $tanggal;
        $this->lokasi  = $lokasi;
        $this->data    = $data;
    }

    /* =====================================================
     | BASIC
     ===================================================== */
    public function collection()
    {
        return collect([]);
    }

    public function title(): string
    {
        return '01';
    }

    /* =====================================================
     | EVENTS
     ===================================================== */
    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $sheet = $event->sheet->getDelegate();

                $this->applyGlobalStyle($sheet);

                // buat objek image
                $drawing = new Drawing();
                $drawing->setName('Image');
                $drawing->setDescription('Image N2:N4');
                $drawing->setPath(public_path('images/pertamina-logo.png'));

                // ukuran image (pixel)
                $drawing->setWidth(112);

                // anchor ke cell hasil merge
                $drawing->setCoordinates('N2');

                // optional: center-kan di merge cell
                $drawing->setOffsetX(10);
                $drawing->setOffsetY(10);

                $drawing->setWorksheet($sheet);

                $this->renderHeader($sheet);
                $this->renderFooter($sheet);
                $this->renderGpsTable($sheet);
                $this->renderDashcamTable($sheet);

                $gpsFailures = [
                    [
                        'plate'    => 'B 1234 CD',
                        'gps'      => 1,
                        'dashcam'  => 1,
                        'desc'     => 'MAINTENANCE',
                    ],
                    [
                        'plate'    => 'B 5678 EF',
                        'gps'      => 1,
                        'dashcam'  => 1,
                        'desc'     => 'MAINTENANCE',
                    ],
                ];

                $dashcamFailures = [
                    [
                        'plate'    => 'B 1234 CD',
                        'gps'      => 1,
                        'dashcam'  => 2,
                        'desc'     => 'MAINTENANCE',
                    ],
                    [
                        'plate'    => 'B 5678 EF',
                        'gps'      => 1,
                        'dashcam'  => 1,
                        'desc'     => 'MAINTENANCE',
                    ],
                ];

                $this->renderFailureTable($sheet, [
                    'startCol' => 'B',
                    'endCol'   => 'H',
                    'startRow' => 25,
                    'title'    => 'KEGAGALAN OPERASI MOBIL TANGKI',
                    'data'     => $gpsFailures,
                ]);

                $this->renderFailureTable($sheet, [
                    'startCol' => 'K',
                    'endCol'   => 'Q',
                    'startRow' => 25,
                    'title'    => 'KEGAGALAN OPERASI MOBIL TANGKI',
                    'data'     => $dashcamFailures,
                ]);

                // Render Violation Tables & Charts
                $this->renderViolationSection($sheet);

                // Render Driver Behavior Tables & Charts
                $this->renderDriverBehaviorSection($sheet);

                foreach ([2, 3, 4, 5] as $row) {
                    $sheet->getStyle("B{$row}")
                        ->getAlignment()
                        ->setHorizontal(Alignment::HORIZONTAL_LEFT)
                        ->setVertical(Alignment::VERTICAL_CENTER);
                }

                $sheet->setShowGridlines(false);
            }
        ];
    }

    /* =====================================================
     | HEADER
     ===================================================== */
    private function renderHeader($sheet): void
    {
        $sheet->setCellValue('B2', 'LAMPIRAN 3A');
        $sheet->setCellValue('B3', 'LAPORAN HARIAN OPERASIONAL GPS & DASHCAM');
        $sheet->setCellValue('B4', 'TANGGAL  : ' . $this->tanggal);
        $sheet->setCellValue('B5', 'LOKASI       : ' . $this->lokasi);

        
        $sheet->getStyle('B2')->getFont()->setBold(true)->setSize(14);
        $sheet->getStyle('B3:B5')->getFont()->setBold(true)->setSize(12);
    }

    /* =====================================================
     | FOOTER
     ===================================================== */
     private function renderFooter($sheet): void
     {
        $sheet->setCellValue('K83', 'Disetujui Oleh,');
        $sheet->setCellValue('K84', 'PT Pertamina Patra Niaga');
        $sheet->setCellValue('K85', 'Sr. Supervisor I Fleet');
        
        $sheet->setCellValue('B85', 'Operator');
        $sheet->setCellValue('B93', 'Billy Damas');

        $sheet->setCellValue('K93', 'I Ketut Purna Wiajaya');
        

        
        $sheet->getStyle('B85:K93')->getFont()->setSize(12);
        $sheet->getStyle('B85')->getFont()->setBold(true);
        $sheet->getStyle('K84:K85')->getFont()->setBold(true);
     }

    /* =====================================================
     | GPS TABLE
     ===================================================== */
    private function renderGpsTable($sheet): void
    {
        $sheet->setCellValue('B7', 'GPS Offline');
        $sheet->setCellValue('C7', 'Unit GPS');

        $this->applyHeaderStyle($sheet, 'B7:C7');

        // Calculate percentages
        $total = $this->data['gps_terpasang'];
        $onlinePercent = $total > 0 ? round(($this->data['gps_online'] / $total) * 100) : 0;
        $offlinePercent = $total > 0 ? round(($this->data['gps_offline'] / $total) * 100) : 0;

        $sheet->fromArray([
            ['Jumlah MT Terpasang GPS', $this->data['gps_terpasang'], ''],
            ['Online', $this->data['gps_online'], ''],
            ['Offline', $this->data['gps_offline'], ''],
            ['% MT Online GPS', $this->data['gps_online_persen'] . '%', ''],
        ], null, 'B8');

        $this->applyBodyStyle($sheet, 'B8:C11', Alignment::HORIZONTAL_LEFT);
        $this->applyHorizontalAlignment($sheet, 'B8:B11', Alignment::HORIZONTAL_LEFT);
        $this->applyHorizontalAlignment($sheet, 'C8:C11', Alignment::HORIZONTAL_RIGHT);
    }

    /* =====================================================
     | DASHCAM TABLE
     ===================================================== */
    private function renderDashcamTable($sheet): void
    {
        $sheet->setCellValue('E7', 'Dashcam Offline');
        $sheet->setCellValue('F7', 'Unit GPS');

        $this->applyHeaderStyle($sheet, 'E7:F7');

        // Calculate percentages
        $total = $this->data['dashcam_terpasang'];
        $onlinePercent = $total > 0 ? round(($this->data['dashcam_online'] / $total) * 100) : 0;
        $offlinePercent = $total > 0 ? round(($this->data['dashcam_offline'] / $total) * 100) : 0;

        $sheet->fromArray([
            ['Jumlah MT Terpasang Dashcam', $this->data['dashcam_terpasang'], ''],
            ['Online', $this->data['dashcam_online'], ''],
            ['Offline', $this->data['dashcam_offline'], ''],
            ['% MT Online GPS', $this->data['dashcam_online_persen'] . '%', ''],
        ], null, 'E8');

        $this->applyBodyStyle($sheet, 'E8:F11', Alignment::HORIZONTAL_LEFT);
        $this->applyHorizontalAlignment($sheet, 'F8:F11', Alignment::HORIZONTAL_RIGHT);
    }

    /* =====================================================
     | FAILURE TABLE (REUSABLE)
     ===================================================== */
    private function renderFailureTable($sheet, array $config): void
    {
        $startCol = $config['startCol'];
        $endCol   = $config['endCol'];
        $startRow = $config['startRow'];
        $title    = $config['title'] ?? 'KEGAGALAN OPERASI MOBIL TANGKI';
        $data     = $config['data'] ?? [];
        $minRows  = 13; // Minimum 13 baris body

        $header1   = $startRow + 1;
        $header2   = $startRow + 2;
        $bodyStart = $startRow + 3;
        $bodyEnd   = $bodyStart + $minRows - 1; // Selalu 13 baris
        $footer1   = $bodyEnd + 1;
        $footer2   = $footer1 + 1;

        // ================= TITLE =================
        $sheet->mergeCells("{$startCol}{$startRow}:{$endCol}{$startRow}");
        $sheet->setCellValue("{$startCol}{$startRow}", $title);
        $sheet->getStyle("{$startCol}{$startRow}")->getFont()->setBold(true)->setSize(12);
        $this->applyHorizontalAlignment($sheet, "{$startCol}{$startRow}", Alignment::HORIZONTAL_CENTER);

        // ================= HEADER =================
        $colNopol   = $startCol;
        $colGps     = chr(ord($startCol) + 1);
        $colDashcam = chr(ord($startCol) + 3);
        $colDesc    = chr(ord($startCol) + 4);

        $sheet->mergeCells("{$colNopol}{$header1}:{$colNopol}".($header1 + 1));
        $sheet->mergeCells("{$colGps}{$header1}:".chr(ord($colGps) + 2).$header1);
        $sheet->mergeCells("{$colDesc}{$header1}:{$endCol}".($header1 + 1));

        $sheet->setCellValue("{$colNopol}{$header1}", 'Nopol');
        $sheet->setCellValue("{$colGps}{$header1}", 'Status');
        $sheet->setCellValue("{$colDesc}{$header1}", 'Keterangan');

        $sheet->mergeCells("{$colGps}{$header2}:".chr(ord($colGps) + 1).$header2);
        $sheet->setCellValue("{$colGps}{$header2}", 'GPS');
        $sheet->setCellValue("{$colDashcam}{$header2}", 'Dashcam');

        $this->applyHeaderStyle($sheet, "{$startCol}{$header1}:{$endCol}{$header2}");

        // ================= BODY =================
        $row = $bodyStart;

        // Render data yang ada
        for ($i = 0; $i < $minRows; $i++) {
            $currentRow = $bodyStart + $i;
            
            if (isset($data[$i])) {
                // Ada data, render data asli
                $item = $data[$i];
                $sheet->setCellValue("{$colNopol}{$currentRow}", $item['plate'] ?? '-');

                $sheet->mergeCells("{$colGps}{$currentRow}:".chr(ord($colGps) + 1).$currentRow);
                $sheet->setCellValue("{$colGps}{$currentRow}", $item['gps'] == 1 ? 'Online' : 'Offline');
                $this->applyHorizontalAlignment($sheet, "{$colGps}{$currentRow}", Alignment::HORIZONTAL_CENTER);

                $sheet->setCellValue("{$colDashcam}{$currentRow}", $item['dashcam'] == 1 ? 'Online' : 'Offline');
                $this->applyHorizontalAlignment($sheet, "{$colDashcam}{$currentRow}", Alignment::HORIZONTAL_CENTER);

                $sheet->mergeCells("{$colDesc}{$currentRow}:".chr(ord($colDesc) + 2).$currentRow);
                $sheet->setCellValue("{$colDesc}{$currentRow}", $item['desc'] ?? '-');
            } else {
                // Tidak ada data, render baris kosong
                $sheet->setCellValue("{$colNopol}{$currentRow}", '');
                
                $sheet->mergeCells("{$colGps}{$currentRow}:".chr(ord($colGps) + 1).$currentRow);
                $sheet->setCellValue("{$colGps}{$currentRow}", '');
                
                $sheet->setCellValue("{$colDashcam}{$currentRow}", '');
                
                $sheet->mergeCells("{$colDesc}{$currentRow}:".chr(ord($colDesc) + 2).$currentRow);
                $sheet->setCellValue("{$colDesc}{$currentRow}", '');
            }
        }

        // Apply border ke semua 13 baris
        $this->applyBorder($sheet, "{$startCol}{$bodyStart}:{$endCol}{$bodyEnd}");

        // ================= FOOTER =================
        $sheet->setCellValue("{$startCol}{$footer1}", 'Jumlah MT');

        $sheet->mergeCells(
            $this->col($startCol, 1) . $footer1 . ':' .
            $this->col($startCol, 2) . $footer1
        );
        $sheet->setCellValue(
            $this->col($startCol, 1) . $footer1,
            'GPS NOT ACTIVE'
        );

        $sheet->setCellValue(
            $this->col($startCol, 3) . $footer1,
            'Dashcam NOT ACTIVE'
        );

        $sheet->mergeCells(
            $this->col($startCol, 4) . $footer1 . ':' .
            $this->col($startCol, 6) . $footer1
        );

        $sheet->setCellValue("{$startCol}{$footer2}", count($data));

        $sheet->mergeCells(
            $this->col($startCol, 1) . $footer2 . ':' .
            $this->col($startCol, 2) . $footer2
        );
        $sheet->setCellValue(
            $this->col($startCol, 1) . $footer2,
            10
        );

        $sheet->setCellValue(
            $this->col($startCol, 3) . $footer2,
            14
        );

        $sheet->mergeCells(
            $this->col($startCol, 4) . $footer2 . ':' .
            $this->col($startCol, 6) . $footer2
        );

        $this->applyFooterStyle($sheet, "{$startCol}{$footer1}:{$endCol}{$footer2}");
    }

    /* =====================================================
     | VIOLATION SECTION (NEW)
     ===================================================== */
    private function renderViolationSection($sheet): void
    {
        $startRow = 44; // last col failures + 2

        // Data pelanggaran
        $violations = [
            'harsh_cornering' => [
                ['region' => 'Region 3 JBB', 'case' => 14],
                ['region' => 'Region 4 JBT', 'case' => 10],
                ['region' => 'Region 5 Jatimbalinus', 'case' => 8],
                ['region' => 'Total', 'case' => 32],
            ],
            'over_speed' => [
                ['region' => 'Region 3 JBB', 'case' => 8],
                ['region' => 'Region 4 JBT', 'case' => 10],
                ['region' => 'Region 5 Jatimbalinus', 'case' => 20],
                ['region' => 'Total', 'case' => 38],
            ],
            'harsh_brake' => [
                ['region' => 'Region 3 JBB', 'case' => 10],
                ['region' => 'Region 4 JBT', 'case' => 5],
                ['region' => 'Region 5 Jatimbalinus', 'case' => 15],
                ['region' => 'Total', 'case' => 30],
            ],
            'harsh_acceleration' => [
                ['region' => 'Region 3 JBB', 'case' => '[ha_jbb]'],
                ['region' => 'Region 4 JBT', 'case' => '[ha_jbt]'],
                ['region' => 'Region 5 Jatimbalinus', 'case' => '[ha_jatimbalinus]'],
                ['region' => 'Total', 'case' => 0],
            ],
            'max_parking' => [
                ['region' => 'Region 3 JBB', 'case' => '[mp_jbb]'],
                ['region' => 'Region 4 JBT', 'case' => '[mp_jbt]'],
                ['region' => 'Region 5 Jatimbalinus', 'case' => '[mp_jatimbalinus]'],
                ['region' => 'Total', 'case' => 0],
            ],
        ];

        // Render Harsh Cornering
        $this->renderViolationTable($sheet, 'B', $startRow, 'Harsh Cornering', $violations['harsh_cornering']);
        
        // Render Over Speed
        $this->renderViolationTable($sheet, 'E', $startRow, 'Over Speed', $violations['over_speed']);
        
        // Render Harsh Brake
        $this->renderViolationTable($sheet, 'H', $startRow, 'Harsh Brake', $violations['harsh_brake']);
        
        // Render Harsh Acceleration
        $this->renderViolationTable($sheet, 'K', $startRow, 'Harsh Acceleration', $violations['harsh_acceleration']);
        
        // Render Max Parking
        $this->renderViolationTable($sheet, 'N', $startRow, 'Max Parking', $violations['max_parking']);
    }

    /* =====================================================
     | DRIVER BEHAVIOR SECTION (NEW)
     ===================================================== */
    private function renderDriverBehaviorSection($sheet): void
    {
        $startRow = 63; // After violation section + space

        // Data driver behavior
        $behaviors = [
            'blackzone' => [
                ['region' => 'Region 3 JBB', 'case' => 10],
                ['region' => 'Region 4 JBT', 'case' => 20],
                ['region' => 'Region 5 Jatimbalinus', 'case' => 15],
                ['region' => 'Total', 'case' => 45],
            ],
            'menguap' => [
                ['region' => 'Region 3 JBB', 'case' => 30],
                ['region' => 'Region 4 JBT', 'case' => 10],
                ['region' => 'Region 5 Jatimbalinus', 'case' => 8],
                ['region' => 'Total', 'case' => 48],
            ],
            'merokok' => [
                ['region' => 'Region 3 JBB', 'case' => '[dbp_jbb]'],
                ['region' => 'Region 4 JBT', 'case' => '[dbp_jbt]'],
                ['region' => 'Region 5 Jatimbalinus', 'case' => '[dbp_jatimbalinus]'],
                ['region' => 'Total', 'case' => 0],
            ],
            'telepon' => [
                ['region' => 'Region 3 JBB', 'case' => '[dbt_jbb]'],
                ['region' => 'Region 4 JBT', 'case' => '[dbt_jbt]'],
                ['region' => 'Region 5 Jatimbalinus', 'case' => '[dbt_jatimbalin]'],
                ['region' => 'Total', 'case' => 0],
            ],
            'distract' => [
                ['region' => 'Region 3 JBB', 'case' => '[dbd_jbb]'],
                ['region' => 'Region 4 JBT', 'case' => '[dbd_jbt]'],
                ['region' => 'Region 5 Jatimbalinus', 'case' => '[dbd_jatimbalin]'],
                ['region' => 'Total', 'case' => 0],
            ],
        ];

        // Render Blackzone
        $this->renderViolationTable($sheet, 'B', $startRow, 'Blackzone', $behaviors['blackzone']);
        
        // Render Driver Behavior - Menguap
        $this->renderViolationTable($sheet, 'E', $startRow, 'Driver Behavior - Menguap', $behaviors['menguap']);
        
        // Render Driver Behavior - Merokok
        $this->renderViolationTable($sheet, 'H', $startRow, 'Driver Behavior - Merokok', $behaviors['merokok']);
        
        // Render Driver Behavior - Telepon
        $this->renderViolationTable($sheet, 'K', $startRow, 'Driver Behavior - Telepon', $behaviors['telepon']);
        
        // Render Driver Behavior - Distract
        $this->renderViolationTable($sheet, 'N', $startRow, 'Driver Behavior - Distract', $behaviors['distract']);
    }

    private function renderViolationTable($sheet, string $startCol, int $startRow, string $title, array $data): void
    {
        $endCol = chr(ord($startCol) + 1);
        
        // Header
        $headerRow = $startRow + 1;
        $sheet->setCellValue("{$startCol}{$headerRow}", $title);
        $sheet->setCellValue("{$endCol}{$headerRow}", 'Case');
        $this->applyHeaderStyle($sheet, "{$startCol}{$headerRow}:{$endCol}{$headerRow}");

        // Body
        $row = $headerRow + 1;
        foreach ($data as $item) {
            $sheet->setCellValue("{$startCol}{$row}", $item['region']);
            $sheet->setCellValue("{$endCol}{$row}", $item['case']);
            
            // Apply style
            if ($item['region'] === 'Total') {
                $sheet->getStyle("{$startCol}{$row}:{$endCol}{$row}")->applyFromArray([
                    'font' => ['bold' => true],
                ]);
            }
            
            $this->applyBorder($sheet, "{$startCol}{$row}:{$endCol}{$row}");
            $row++;
        }
    }

    /* =====================================================
     | STYLES
     ===================================================== */
    private function applyHeaderStyle($sheet, string $range): void
    {
        $sheet->getStyle($range)->applyFromArray([
            'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '262A4A'],
            ],
            'borders' => [
                'allBorders' => ['borderStyle' => Border::BORDER_THIN],
            ],
        ]);
    }

    private function applyBodyStyle($sheet, string $range, string $horizontalAlignment = Alignment::HORIZONTAL_LEFT): void
    {
        $this->applyBorder($sheet, $range);
        $this->applyHorizontalAlignment($sheet, $range, $horizontalAlignment);
    }

    private function applyHorizontalAlignment($sheet, string $range, string $horizontalAlignment = Alignment::HORIZONTAL_LEFT): void
    {
        $sheet->getStyle($range)->getAlignment()->setHorizontal($horizontalAlignment);
    }

    private function applyFooterStyle($sheet, string $range): void
    {
        $sheet->getStyle($range)->applyFromArray([
            'font' => ['bold' => true],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '9EC3F7'],
            ],
            'borders' => [
                'allBorders' => ['borderStyle' => Border::BORDER_THIN],
            ],
        ]);
    }

    private function applyBorder($sheet, string $range): void
    {
        $sheet->getStyle($range)->getBorders()->getAllBorders()
            ->setBorderStyle(Border::BORDER_THIN);
    }

    private function applyGlobalStyle($sheet): void
    {
        $sheet->getParent()->getDefaultStyle()
            ->getFont()->setName('Arial');
    }

    private function col(string $col, int $offset = 0): string
    {
        return chr(ord($col) + $offset);
    }

    /* =====================================================
     | CHARTS
     ===================================================== */
    public function charts(): array
    {
        $charts = [
            $this->pieChart('gps', 'GPS Status', 'B9:B10', 'C9:C10', 'B13', 'D23'),
            $this->pieChart('dashcam', 'Dashcam Status', 'E9:E10', 'F9:F10', 'E13', 'G23'),
        ];

        // Add violation charts
        $charts[] = $this->barChart('harsh_cornering', 'Harsh Cornering', 'B46:B49', 'C46:C49', 'B51', 'D61');
        $charts[] = $this->barChart('over_speed', 'Over Speed', 'E46:E49', 'F46:F49', 'E51', 'G61');
        $charts[] = $this->barChart('harsh_brake', 'Harsh Brake', 'H46:H49', 'I46:I49', 'H51', 'J61');
        $charts[] = $this->barChart('harsh_acceleration', 'Harsh Acceleration', 'K46:K49', 'L46:L49', 'K51', 'M61');
        $charts[] = $this->barChart('max_parking', 'Max Parking', 'N46:N49', 'O46:O49', 'N51', 'P61');

        // Add driver behavior charts
        $charts[] = $this->barChart('blackzone', 'Blackzone', 'B65:B68', 'C65:C68', 'B70', 'D80');
        $charts[] = $this->barChart('menguap', 'Driver Behavior - Menguap', 'E65:E68', 'F65:F68', 'E70', 'G80');
        $charts[] = $this->barChart('merokok', 'Driver Behavior - Merokok', 'H65:H68', 'I65:I68', 'H70', 'J80');
        $charts[] = $this->barChart('telepon', 'Driver Behavior - Telepon', 'K65:K68', 'L65:L68', 'K70', 'M80');
        $charts[] = $this->barChart('distract', 'Driver Behavior - Distract', 'N65:N68', 'O65:O68', 'N70', 'P80');

        return $charts;
    }

    private function pieChart(
        string $id,
        string $title,
        string $categoryRange,
        string $valueRange,
        string $topLeft,
        string $bottomRight
    ): Chart {

        $categories = new DataSeriesValues(
            DataSeriesValues::DATASERIES_TYPE_STRING,
            "'01'!\${$categoryRange}",
            null,
            2
        );

        $values = new DataSeriesValues(
            DataSeriesValues::DATASERIES_TYPE_NUMBER,
            "'01'!\${$valueRange}",
            null,
            2
        );

        $series = new DataSeries(
            DataSeries::TYPE_PIECHART,
            null,
            [0],
            [],
            [$categories],
            [$values]
        );

        // Try to enable data labels (may not work in all PhpSpreadsheet versions)
        try {
            $layout = new \PhpOffice\PhpSpreadsheet\Chart\Layout();
            // $layout->setShowVal(true);
            $layout->setShowPercent(true);
            $plotArea = new PlotArea($layout, [$series]);
        } catch (\Exception $e) {
            // Fallback to basic chart without custom layout
            $plotArea = new PlotArea(null, [$series]);
        }

        $legend   = new Legend(Legend::POSITION_RIGHT);
        $titleObj = new Title($title);

        $chart = new Chart(
            $id,
            $titleObj,
            $legend,
            $plotArea
        );

        $chart->setTopLeftPosition($topLeft);
        $chart->setBottomRightPosition($bottomRight, 0);

        return $chart;
    }

    private function barChart(
        string $id,
        string $title,
        string $categoryRange,
        string $valueRange,
        string $topLeft,
        string $bottomRight
    ): Chart {
        
        $layout = new Layout();
        $layout->setShowVal(true);

        $seriesLabel = [
            new DataSeriesValues(
                DataSeriesValues::DATASERIES_TYPE_STRING,
                "'01'!\$B\$46",
                null,
                1
            ),
        ];
        
        $categories = new DataSeriesValues(
            DataSeriesValues::DATASERIES_TYPE_STRING,
            "'01'!\${$categoryRange}",
            null,
            3
        );

        $values = new DataSeriesValues(
            DataSeriesValues::DATASERIES_TYPE_NUMBER,
            "'01'!\${$valueRange}",
            null,
            3
        );

        $series = new DataSeries(
            DataSeries::TYPE_BARCHART,
            DataSeries::GROUPING_CLUSTERED,
            [0],
            $seriesLabel,
            [$categories],
            [$values]
        );

        $series->setPlotDirection(DataSeries::DIRECTION_COL);

        $richText = new RichText();
        $run1 = $richText->createTextRun($title);
        $run1->getFont()->setBold(true)->setSize(12);

        $plotArea = new PlotArea($layout, [$series]);
        $legend   = new Legend(Legend::POSITION_RIGHT);
        $titleObj = new Title($richText);

        $chart = new Chart(
            $id,
            $titleObj,
            $legend,
            $plotArea
        );

        $chart->setTopLeftPosition($topLeft);
        $chart->setBottomRightPosition($bottomRight, 0);

        return $chart;
    }

    /* =====================================================
     | COLUMN WIDTHS
     ===================================================== */
    public function columnWidths(): array
    {
        return [
            'A' => 3.33, 'B' => 26.17, 'C' => 13.67, 'D' => 3.33,
            'E' => 26.17, 'F' => 14.50, 'G' => 3.33, 'H' => 26.17,
            'I' => 14.67, 'J' => 3.33, 'K' => 26.17, 'L' => 13.67,
            'M' => 3.33, 'N' => 22, 'O' => 14.17, 'P' => 1, 'Q' => 1.67,
        ];
    }
}