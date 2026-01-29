<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class KpiExport implements FromCollection, WithEvents, WithColumnWidths, WithTitle
{
    protected array $data;
    protected string $sheetName;

    public function __construct(array $data, string $sheetName = '01')
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

                // Left Logo
                // buat objek image
                $drawing = new Drawing();
                $drawing->setName('Image');
                $drawing->setDescription('Image B1:B3');
                $drawing->setPath(public_path('ids.png'));

                // ukuran image (pixel)
                $drawing->setWidth(180);

                // anchor ke cell hasil merge
                $drawing->setCoordinates('B1');

                // optional: center-kan di merge cell
                $drawing->setOffsetX(0);
                $drawing->setOffsetY(0);

                $drawing->setWorksheet($sheet);

                // Right Logo
                // buat objek image
                $rightLogo = new Drawing();
                $rightLogo->setName('Image');
                $rightLogo->setDescription('Image K1:L5');
                $rightLogo->setPath(public_path('pertamina-patra-logistik.png'));

                // ukuran image (pixel)
                $rightLogo->setWidth(180);

                // anchor ke cell hasil merge
                $rightLogo->setCoordinates("K1");

                // optional: center-kan di merge cell
                $rightLogo->setOffsetX(70);
                $rightLogo->setOffsetY(15);

                $rightLogo->setWorksheet($sheet);

                $this->renderHeader($sheet);
                $this->renderItemTable($sheet);
                $this->renderFooter($sheet);

                foreach ([5, 6, 7, 8] as $row) {
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
        $sheet->setCellValue('B5', 'KEY PERFORMANCE INDICATOR (KPI)');
        $sheet->setCellValue('B6', 'JASA SEWA PERANGKAT GPS DAN DASHCAM');
        $sheet->setCellValue('B7', 'PERIODE    : SEPTEMBER 2025');
        $sheet->setCellValue('B8', 'LOKASI       : IT Jakarta');
        
        $sheet->getStyle('B5:B8')->getFont()->setBold(true)->setSize(12)->setName("Calibri");
    }

    /* =====================================================
     | FOOTER
     ===================================================== */
     private function renderFooter($sheet): void
     {
        $sheet->setCellValue('B28', 'Disiapkan Oleh,');
        $sheet->setCellValue('B29', 'PT Indi Daya Sistem');
        $sheet->setCellValue('B30', 'Supervisor Services');
        $sheet->setCellValue('B37', 'Dimas Nafidin');

        $lastCol = 'H';
        $sheet->setCellValue("{$lastCol}28", 'Disetujui Oleh,');
        $sheet->setCellValue("{$lastCol}29", 'PT  Patra Logistik');
        $sheet->setCellValue("{$lastCol}30", 'Area Manager Jawa Bagian Barat');
        $sheet->setCellValue("{$lastCol}37", 'Bayu Riyadi');
        
        $sheet->getStyle("B28:{$lastCol}37")->getFont()->setSize(14);
        $sheet->getStyle("B29:B37")->getFont()->setBold(true);
        $sheet->getStyle("{$lastCol}29:{$lastCol}37")->getFont()->setBold(true);
     }

     /* =====================================================
     | TABLE
     ===================================================== */
    private function renderItemTable($sheet): void
    {
        $headerRow1 = 11;
        $headerRow2 = $headerRow1 + 1;
        $headerRow3 = $headerRow2 + 1;

        $column1 = 'B';                                      // INDIKATOR KINERJA UTAMA
        $column2 = $this->col($column1, 3);     // POLARITY
        $column3 = $this->col($column2, 1);     // FREKUENSI MONITORING
        $column4 = $this->col($column3, 1);     // BOBOT (%)
        $column5 = $this->col($column4, 1);     // SATUAN
        $column6 = $this->col($column5, 1);     // TARGET
        $column7 = $this->col($column6, 1);     // REALISASI
        $column8 = $this->col($column7, 1);     // WEIGHTED SCORE
        $column9 = $this->col($column8, 1);     // KETERANGAN

        // HEADER
        $sheet->getRowDimension($headerRow1)->setRowHeight(21);
        $sheet->getRowDimension($headerRow2)->setRowHeight(29);

        $sheet->mergeCells("{$column1}{$headerRow1}:{$this->col($column1, 2)}{$headerRow2}");
        $sheet->mergeCells("{$column2}{$headerRow1}:{$column2}{$headerRow2}");
        $sheet->mergeCells("{$column3}{$headerRow1}:{$column3}{$headerRow2}");
        $sheet->mergeCells("{$column4}{$headerRow1}:{$column4}{$headerRow2}");
        $sheet->mergeCells("{$column5}{$headerRow1}:{$column5}{$headerRow2}");
        $sheet->mergeCells("{$column6}{$headerRow1}:{$column6}{$headerRow2}");

        $sheet->mergeCells("{$column7}{$headerRow1}:{$column8}{$headerRow1}");

        // $sheet->mergeCells("{$column9}{$headerRow1}:{$column9}{$headerRow2}");

        $sheet->setCellValue("{$column1}{$headerRow1}", "INDIKATOR KINERJA UTAMA");
        $sheet->setCellValue("{$column2}{$headerRow1}", "POLARITY");
        $sheet->setCellValue("{$column3}{$headerRow1}", "FREKUENSI MONITORING");
        $sheet->setCellValue("{$column4}{$headerRow1}", "BOBOT\n(%)");
        $sheet->setCellValue("{$column5}{$headerRow1}", "SATUAN");
        $sheet->setCellValue("{$column6}{$headerRow1}", "TARGET");
        $sheet->setCellValue("{$column7}{$headerRow1}", "REALISASI");
        $sheet->setCellValue("{$column7}{$headerRow2}", "REALISASI");
        $sheet->setCellValue("{$column8}{$headerRow2}", "WEIGHTED SCORE");
        $sheet->setCellValue("{$column9}{$headerRow2}", "KETERANGAN");

        $sheet->setCellValue("{$column4}{$headerRow3}", "( a )");
        $sheet->setCellValue("{$column6}{$headerRow3}", "( b )");
        $sheet->setCellValue("{$column7}{$headerRow3}", "( c )");
        $sheet->setCellValue("{$column8}{$headerRow3}", "(d) = (c) / (b) x (a)");

        $this->applyHeaderStyle($sheet, "{$column1}{$headerRow1}:{$column9}{$headerRow3}");

        // BODY
        $bodyStartRow   = 14;
        $bodyEndRow     = 23;
        $bodyStartCol   = 'C';
        $bodyEndCol     = 'L';
        $footerRow      = $bodyEndRow + 1;

        $lastRowMain    = $bodyEndRow - 2;

        // Main Metrics
        $sheet->mergeCells("{$column1}{$bodyStartRow}:{$column1}{$lastRowMain}");
        $sheet->setCellValue("{$column1}{$bodyStartRow}", "Main Metrics");
        $sheet->getStyle("{$column1}{$bodyStartRow}")->applyFromArray([
            'font' => ['bold' => true, 'color' => ['rgb' => 'ffffff'], 'size' => 10],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'c00000'],
            ]
        ]);

        
        // Boundary KPI
        $sheet->mergeCells("{$column1}22:{$column1}{$bodyEndRow}");
        $sheet->setCellValue("{$column1}22", "Boundary KPI");
        $sheet->getStyle("{$column1}22")->applyFromArray([
            'font' => ['bold' => true, 'color' => ['rgb' => 'ffffff'], 'size' => 10],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '1d12f6'],
            ]
        ]);

        $columnIndicator = $this->col($bodyStartCol, 1);

        $subRow1 = $bodyStartRow + 1;
        $subRow2 = $subRow1 + 1;
        $subRow3 = $subRow2 + 1;
        $subRow4 = $subRow3 + 1;

        $subRow5 = $subRow4 + 2;
        $subRow6 = $subRow5 + 1;

        $subRow7 = $subRow6 + 2;

        // I. Operational
        $sheet->setCellValue("{$bodyStartCol}{$bodyStartRow}", "I. Operational");
        $sheet->setCellValue("{$column4}{$bodyStartRow}", 70);
        $sheet->setCellValue("{$column8}{$bodyStartRow}", 66.7);
        $this->applyTextColor($sheet, "{$column8}{$bodyStartRow}", "cc0100");
        $this->applyMainRowStyle($sheet, "{$bodyStartCol}{$bodyStartRow}:{$column9}{$bodyStartRow}");

        // Row 1
        $sheet->setCellValue("{$bodyStartCol}{$subRow1}", "1");
        $sheet->setCellValue("{$columnIndicator}{$subRow1}", "Kegagalan Operasi Perangkat");
        $sheet->setCellValue("{$column2}{$subRow1}", "Max");
        $sheet->setCellValue("{$column3}{$subRow1}", "Harian - Bulanan");
        $sheet->setCellValue("{$column4}{$subRow1}", 20);
        $sheet->setCellValue("{$column5}{$subRow1}", "%");
        $sheet->setCellValue("{$column6}{$subRow1}", 100);
        $sheet->setCellValue("{$column7}{$subRow1}", 84.3);
        $sheet->setCellValue("{$column8}{$subRow1}", 16.9);
        $sheet->setCellValue("{$column9}{$subRow1}", "Berdasarkan Lampiran 1&2");

        // Row 2
        $sheet->setCellValue("{$bodyStartCol}{$subRow2}", "2");
        $sheet->setCellValue("{$columnIndicator}{$subRow2}", "Penormalan Operasi Perangkat");
        $sheet->setCellValue("{$column2}{$subRow2}", "Max");
        $sheet->setCellValue("{$column3}{$subRow2}", "Harian - Bulanan");
        $sheet->setCellValue("{$column4}{$subRow2}", 20);
        $sheet->setCellValue("{$column5}{$subRow2}", "%");
        $sheet->setCellValue("{$column6}{$subRow2}", 100);
        $sheet->setCellValue("{$column7}{$subRow2}", 84.3);
        $sheet->setCellValue("{$column8}{$subRow2}", 16.9);
        $sheet->setCellValue("{$column9}{$subRow2}", "Berdasarkan Lampiran 1&2");

        // Row 3
        $sheet->setCellValue("{$bodyStartCol}{$subRow3}", "3");
        $sheet->setCellValue("{$columnIndicator}{$subRow3}", "Timely Performance Reporting");
        $sheet->setCellValue("{$column2}{$subRow3}", "Max");
        $sheet->setCellValue("{$column3}{$subRow3}", "Harian - Bulanan");
        $sheet->setCellValue("{$column4}{$subRow3}", 20);
        $sheet->setCellValue("{$column5}{$subRow3}", "%");
        $sheet->setCellValue("{$column6}{$subRow3}", 100);
        $sheet->setCellValue("{$column7}{$subRow3}", 84.3);
        $sheet->setCellValue("{$column8}{$subRow3}", 16.9);
        $sheet->setCellValue("{$column9}{$subRow3}", "Berdasarkan Lampiran 3c");

        // Row 4
        $sheet->setCellValue("{$bodyStartCol}{$subRow4}", "4");
        $sheet->setCellValue("{$columnIndicator}{$subRow4}", "Response Time Keluhan User/Stakeholder");
        $sheet->setCellValue("{$column2}{$subRow4}", "Max");
        $sheet->setCellValue("{$column3}{$subRow4}", "Harian - Bulanan");
        $sheet->setCellValue("{$column4}{$subRow4}", 20);
        $sheet->setCellValue("{$column5}{$subRow4}", "%");
        $sheet->setCellValue("{$column6}{$subRow4}", 100);
        $sheet->setCellValue("{$column7}{$subRow4}", 84.3);
        $sheet->setCellValue("{$column8}{$subRow4}", 16.9);
        $sheet->setCellValue("{$column9}{$subRow4}", "Berdasarkan Lampiran 6");

        // II. People Management & Audit
        $bodyMainRow2 = 19;
        $sheet->setCellValue("{$bodyStartCol}{$bodyMainRow2}", "II. People Management & Audit");
        $sheet->setCellValue("{$column4}{$bodyMainRow2}", 30);
        $sheet->setCellValue("{$column8}{$bodyMainRow2}", 28.7);
        $this->applyTextColor($sheet, "{$column8}{$bodyMainRow2}", "cc0100");
        $this->applyMainRowStyle($sheet, "{$bodyStartCol}{$bodyMainRow2}:{$column9}{$bodyMainRow2}");
        
        // Row 5
        $sheet->setCellValue("{$bodyStartCol}{$subRow5}", "5");
        $sheet->setCellValue("{$columnIndicator}{$subRow5}", "Tingkat Kehadiran Personel");
        $sheet->setCellValue("{$column2}{$subRow5}", "Max");
        $sheet->setCellValue("{$column3}{$subRow5}", "Harian - Bulanan");
        $sheet->setCellValue("{$column4}{$subRow5}", 15);
        $sheet->setCellValue("{$column5}{$subRow5}", "%");
        $sheet->setCellValue("{$column6}{$subRow5}", 100);
        $sheet->setCellValue("{$column7}{$subRow5}", 91.6);
        $sheet->setCellValue("{$column8}{$subRow5}", 13.7);
        $sheet->setCellValue("{$column9}{$subRow5}", "Berdasarkan Lampiran 5");

        // Row 6
        $sheet->setCellValue("{$bodyStartCol}{$subRow6}", "6");
        $sheet->setCellValue("{$columnIndicator}{$subRow6}", "Tindak Lanjut Temuan MWT, Inspeksi, & Audit");
        $sheet->setCellValue("{$column2}{$subRow6}", "Max");
        $sheet->setCellValue("{$column3}{$subRow6}", "Harian - Bulanan");
        $sheet->setCellValue("{$column4}{$subRow6}", 15);
        $sheet->setCellValue("{$column5}{$subRow6}", "%");
        $sheet->setCellValue("{$column6}{$subRow6}", 100);
        $sheet->setCellValue("{$column7}{$subRow6}", 100);
        $sheet->setCellValue("{$column8}{$subRow6}", 15);
        $sheet->setCellValue("{$column9}{$subRow6}", null);

        // III. Boundary KPI
        $bodyMainRow3 = 22;
        $sheet->setCellValue("{$bodyStartCol}{$bodyMainRow3}", "III. Boundary KPI");
        $this->applyTextColor($sheet, "{$column8}{$bodyMainRow3}", "ff0000");
        $this->applyMainRowStyle($sheet, "{$bodyStartCol}{$bodyMainRow3}:{$column9}{$bodyMainRow3}");

        // Row 7
        $sheet->setCellValue("{$bodyStartCol}{$subRow7}", "7");
        $sheet->setCellValue("{$columnIndicator}{$subRow7}", "Number of Accident (NoA)");
        $sheet->setCellValue("{$column2}{$subRow7}", "Min");
        $sheet->setCellValue("{$column3}{$subRow7}", "Bulanan");
        $sheet->setCellValue("{$column4}{$subRow7}", null);
        $sheet->setCellValue("{$column5}{$subRow7}", "Kasus");
        $sheet->setCellValue("{$column6}{$subRow7}", "-");
        $sheet->setCellValue("{$column7}{$subRow7}", null);
        $sheet->setCellValue("{$column8}{$subRow7}", null);
        $sheet->setCellValue("{$column9}{$subRow7}", null);

        // FOOTER
        $sheet->setCellValue("D{$footerRow}", "GRAND TOTAL");
        
        $sheet->setCellValue("{$column8}{$footerRow}", 95.5);
        $this->applyTextColor($sheet, "{$column8}{$footerRow}", "ff0000");
        $sheet->getStyle("{$column1}{$footerRow}:{$column8}{$footerRow}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);


        // Styling Body
        $sheet->getStyle("{$column4}{$bodyStartRow}:{$column4}{$subRow7}")->getNumberFormat()->setFormatCode('#,##0"%"');


        $sheet->getStyle("{$column2}{$bodyStartRow}:{$column7}{$subRow7}")
            ->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->getStyle("{$bodyStartCol}{$bodyStartRow}:{$column8}{$subRow7}")
            ->getFont()->setSize(10);

        $sheet->getStyle("{$column9}{$bodyStartRow}:{$column9}{$subRow7}")
            ->getFont()->setSize(8);

        $sheet->getStyle("{$column1}{$footerRow}:{$column9}{$footerRow}")
            ->getFont()->setBold(true)->setSize(12);

        
        $sheet->getStyle("{$column1}{$bodyStartRow}:{$column1}{$subRow7}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle("{$column2}{$bodyStartRow}:{$column2}{$subRow7}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle("{$column3}{$bodyStartRow}:{$column3}{$subRow7}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle("{$column4}{$bodyStartRow}:{$column4}{$footerRow}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle("{$column5}{$bodyStartRow}:{$column5}{$subRow7}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle("{$column6}{$bodyStartRow}:{$column6}{$subRow7}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle("{$column7}{$bodyStartRow}:{$column7}{$subRow7}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle("{$column8}{$bodyStartRow}:{$column8}{$subRow7}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle("{$column9}{$bodyStartRow}:{$column9}{$footerRow}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_THIN);
    }

    /* =====================================================
     | STYLES
     ===================================================== */
    private function applyMainRowStyle($sheet, string $range): void
    {
        $this->applyBackgroundColor($sheet, $range, "c7d9f0");
        $sheet->getStyle($range)->getFont()->setBold(true);
    }
    private function applyTextColor($sheet, string $range, string $color)
    {
        $sheet->getStyle($range)->applyFromArray([
            'font' => ['color' => ['rgb' => $color ?? '000000']],
        ]);
    }
    private function applyBackgroundColor($sheet, string $range, string $color)
    {
        $sheet->getStyle($range)->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => $color ?? 'ffffff'],
            ],
        ]);
    }
    private function applyHeaderStyle($sheet, string $range): void
    {
        $sheet->getStyle($range)->applyFromArray([
            'font' => ['bold' => true, 'color' => ['rgb' => 'ffffff'], 'size' => 10],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '002060'],
            ],
            'borders' => [
                'allBorders' => ['borderStyle' => Border::BORDER_THIN],
            ],
        ]);

        $sheet->getStyle($range)->getAlignment()->setWrapText(true);
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
            ->getFont()->setName('Calibri');
    }
    private function col(string $col, int $offset = 0): string
    {
        return chr(ord($col) + $offset);
    }

    /* =====================================================
     | COLUMN WIDTHS
     ===================================================== */
    public function columnWidths(): array
    {
        $data = [
            'A' => 3.67, 'B' => 11.33, 'C' => 3.00, 'D' => 36.17, 'E' => 9.33, 'F' => 12.17, 'G' => 9.33,
            'H' => 7, 'I' => 8.67, 'J' => 9.67, 'K' => 15.67, 'L' => 15.67
        ];

        return array_map(fn($v) => $v + 0.83, $data);
    }
}