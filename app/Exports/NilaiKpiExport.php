<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class NilaiKpiExport implements FromCollection, WithEvents, WithColumnWidths, WithTitle
{
    protected array $data;
    protected string $sheetName;

    public function __construct(array $data, string $sheetName = 'Nilai KPI')
    {
        $this->data = $data;
        $this->sheetName = $sheetName;
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
        return $this->sheetName; // Sheet Name
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

                $kegagalan_operasi = $this->data['kegagalan_operasi'] ?? [];
                $response_time_keluhan = $this->data['response_time_keluhan'] ?? [];                
                $tindak_lanjut_temuan = $this->data['tindak_lanjut_temuan'] ?? [];                

                $this->renderTable($sheet, [
                    'startCol'  => 'A',
                    'endCol'    => 'B',
                    'startRow'  => 1,
                    'column_1'  => 'Jumlah Kasus',
                    'title'     => '1. Nama KPI: Kegagalan Operasi ',
                    'data'      => $kegagalan_operasi
                ]);

                $this->renderTable($sheet, [
                    'startCol'  => 'D',
                    'endCol'    => 'E',
                    'startRow'  => 1,
                    'title'     => '2. Nama KPI: Penormalan Kegagalan Ops',
                    'data'      => $kegagalan_operasi
                ]);

                $this->renderTable($sheet, [
                    'startCol'  => 'G',
                    'endCol'    => 'H',
                    'startRow'  => 1,
                    'title'     => '4. Nama KPI: Response Time Keluhan User',
                    'data'      => $response_time_keluhan
                ]);

                $this->renderTable($sheet, [
                    'startCol'  => 'J',
                    'endCol'    => 'K',
                    'startRow'  => 1,
                    'title'     => '6. Nama KPI: Tindak Lanjut Temuan',
                    'data'      => $tindak_lanjut_temuan
                ]);

                // $sheet->setShowGridlines(false);
            }
        ];
    }

     /* =====================================================
     | RENDER TABLE
     ===================================================== */
    private function renderTable($sheet, array $config): void
    {
        $startCol = $config['startCol'];
        $endCol   = $config['endCol'];
        $startRow = $config['startRow'];
        $title    = $config['title'] ?? 'Nama KPI';
        $data     = $config['data'] ?? [];
        $minRows  = count($data) > 10 ? count($data) : 10; // Minimum 10 baris body
        
        $headerRow = $startRow + 1;
        $bodyStart = $headerRow + 1;
        $bodyEnd   = $bodyStart + $minRows - 1;
        
        // ================= TITLE =================
        $sheet->setCellValue("{$startCol}{$startRow}", $title);
        $sheet->getStyle("{$startCol}{$startRow}")->getFont()->setBold(true)->setItalic(true)->setSize(11);

        // ================= HEADER =================
        $column1 = $startCol;
        $column2 = chr(ord($startCol) + 1);

        $column1Title = $config['column_1'] ?? 'Waktu';

        $sheet->setCellValue("{$column1}{$headerRow}", $column1Title);
        $sheet->setCellValue("{$column2}{$headerRow}", 'Nilai KPI');

        $this->applyHeaderStyle($sheet, "{$startCol}{$headerRow}:{$endCol}{$headerRow}");

        // ================= BODY =================
        // Render data yang ada
        for ($i = 0; $i < $minRows; $i++) {
            $currentRow = $bodyStart + $i;
            
            if (isset($data[$i])) {
                // Ada data, render data asli
                $item = $data[$i];
                $sheet->setCellValue("{$column1}{$currentRow}", $item['value'] ?? $item['time']);
                $sheet->setCellValue("{$column2}{$currentRow}", $item['percentage'].'%');
                $sheet->getStyle("{$column2}{$currentRow}")->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_PERCENTAGE);

                $this->applyHorizontalAlignment($sheet, "{$column1}{$currentRow}:{$column2}{$currentRow}", Alignment::HORIZONTAL_CENTER);
            } else {
                // Tidak ada data, render baris kosong
                $sheet->setCellValue("{$column1}{$currentRow}", '');
                $sheet->setCellValue("{$column2}{$currentRow}", '');
            }
        }

        // Apply border ke semua 10 baris
        $this->applyBorder($sheet, "{$startCol}{$bodyStart}:{$endCol}{$bodyEnd}");
    }

    /* =====================================================
     | STYLES
     ===================================================== */
     private function applyHeaderStyle($sheet, string $range): void
    {
        $sheet->getStyle($range)->applyFromArray([
            'font' => ['color' => ['rgb' => 'FFFFFF']],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '1f497d'],
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

    private function applyBorder($sheet, string $range): void
    {
        $sheet->getStyle($range)->getBorders()->getAllBorders()
            ->setBorderStyle(Border::BORDER_THIN);
    }

    private function applyGlobalStyle($sheet): void
    {
        $sheet->getParent()->getDefaultStyle()
            ->getFont()->setName('Calibri');
    }

    /* =====================================================
     | COLUMN WIDTHS
     ===================================================== */
    public function columnWidths(): array
    {
        return [
            'A' => 11.16, 'B' => 14.5, 'C' => 12.33, 'D' => 11.66, 'E' => 11.66,
            'F' => 12.5, 'G' => 12.33, 'H' => 12.33, 'I' => 12.33, 'J' => 11.66,
            'K' => 11.66
        ];
    }
}