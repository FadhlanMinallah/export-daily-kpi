<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class TimelyReportingExport implements FromCollection, WithEvents, WithColumnWidths, WithTitle
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

                // buat objek image
                $drawing = new Drawing();
                $drawing->setName('Image');
                $drawing->setDescription('Image N2:N4');
                $drawing->setPath(public_path('images/pertamina-logo.png'));

                // ukuran image (pixel)
                $drawing->setWidth(112);

                // anchor ke cell hasil merge
                $drawing->setCoordinates('AI1');

                // optional: center-kan di merge cell
                $drawing->setOffsetX(10);
                $drawing->setOffsetY(10);

                $drawing->setWorksheet($sheet);

                $this->renderHeader($sheet);
                $this->renderFooter($sheet);
                $this->renderItemTable($sheet);

                foreach ([1, 2, 3, 4] as $row) {
                    $sheet->getStyle("B{$row}")
                        ->getAlignment()
                        ->setHorizontal(Alignment::HORIZONTAL_LEFT)
                        ->setVertical(Alignment::VERTICAL_CENTER);
                }

                $sheet->setShowGridlines(false);

                $sheet->getRowDimension(6)->setRowHeight(26);
            }
        ];
    }

    /* =====================================================
     | HEADER
     ===================================================== */
    private function renderHeader($sheet): void
    {
        $sheet->setCellValue('B1', 'LAMPIRAN 3B');
        $sheet->setCellValue('B2', 'REKAPITULASI LAPORAN BULANAN OPERASIONAL GPS & DASHCAM');
        $sheet->setCellValue('B3', 'PERIODE    : ' . $this->data['periode'] ?? '');
        $sheet->setCellValue('B4', 'LOKASI       : ' . $this->data['site'] ?? '');

        
        $sheet->getStyle('B1')->getFont()->setBold(true)->setSize(14);
        $sheet->getStyle('B2:B4')->getFont()->setBold(true)->setSize(12);
    }

    /* =====================================================
     | FOOTER
     ===================================================== */
     private function renderFooter($sheet): void
     {
        $sheet->setCellValue('B25', 'Disiapkan Oleh,');
        $sheet->setCellValue('AC25', 'Disetujui Oleh,');

        $sheet->setCellValue('B26', 'PT Patra Logistik');
        $sheet->setCellValue('AC26', 'PT Pertamina  Patra Niaga');

        $sheet->setCellValue('AC27', 'Fuel/Integrated Terminal Manager');
        
        $sheet->setCellValue('B32', 'Nama');
        $sheet->setCellValue('AC32', 'Nama');

        
        $sheet->getStyle('B25:AC32')->getFont()->setSize(12);
        $sheet->getStyle('B26:AC32')->getFont()->setBold(true);
     }

     /* =====================================================
     | GPS TABLE
     ===================================================== */
    private function renderItemTable($sheet): void
    {
        // ================= HEADER =================
        // Column "No"
        $sheet->mergeCells('A6:A7');
        $sheet->setCellValue('A6', 'No');

        // Column "Item Check"
        $sheet->mergeCells('B6:B7');
        $sheet->setCellValue('B6', 'Item Check');

        // Column "Periode"
        $sheet->mergeCells('C6:AG6');
        $sheet->setCellValue('C6', '01/01/2026');

        // Column "Date"
        $colPeriode = 'C';
        $rowHeader2 = 7;
        for ($i=1; $i <= 31; $i++) { 
            $sheet->setCellValue($colPeriode.$rowHeader2, $i);
            $colPeriode++;
        }

        // Column "Rata-rata Bulanan"
        $sheet->mergeCells('AH6:AH7');
        $sheet->setCellValue('AH6', 'Rata-rata Bulanan');

        // Column "Keterangan"
        $sheet->mergeCells('AI6:AI7');
        $sheet->setCellValue('AI6', 'Keterangan');

        $sheet->getStyle('AH6')->getAlignment()->setWrapText(true);
        $sheet->getStyle('A6:AI7')->getFont()->setBold(true)->setSize(10);
        $this->applyHeaderStyle($sheet, 'A6:AI7');

        // ================= BODY =================
        $data = [];

        $row = [
            'item_check' => 'Perangkat Status Offline GPS',
        ];

        for ($i = 1; $i <= 31; $i++) {
            $row[(string)$i] = rand(0, 5) . '%';
        }

        $row['avg_month']  = '5%';
        $row['keterangan'] = 'Persentase Perangkat Status Offline GPS';

        $data[] = $row;

        $colStartItemCheck = 'B';
        $rowStartItemCheck = 8;

        foreach ($data as $key => $row) {
            foreach ($row as $key => $column) {
                $sheet->setCellValue($colStartItemCheck.$rowStartItemCheck, $column);
                $colStartItemCheck++;
            }
            $rowStartItemCheck++;
        }

        $this->applyHorizontalAlignment($sheet, 'C8:AH22', Alignment::HORIZONTAL_CENTER);
    }


    /* =====================================================
     | STYLES
     ===================================================== */
    private function applyHeaderStyle($sheet, string $range): void
    {
        $sheet->getStyle($range)->applyFromArray([
            'font' => ['bold' => true, 'color' => ['rgb' => '000000']],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'c1e4f5'],
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
     | COLUMN WIDTHS
     ===================================================== */
    public function columnWidths(): array
    {
        return [
            'A' => 4, 'B' => 43.33, 'C' => 5.33, 'D' => 5.33,
            'E' => 5.33, 'F' => 5.33, 'G' => 5.33, 'H' => 5.33,
            'I' => 5.33, 'J' => 5.33, 'K' => 5.33, 'L' => 5.33,
            'M' => 5.33, 'N' => 5.33, 'O' => 5.33, 'P' => 5.33, 'Q' => 5.33,
            'R' => 5.33, 'S' => 5.33, 'T' => 5.33, 'U' => 5.33, 'V' => 5.33,
            'W' => 5.33, 'X' => 5.33, 'Y' => 5.33, 'Z' => 5.33, 'AA' => 5.33,
            'AB' => 5.33, 'AC' => 5.33, 'AD' => 5.33, 'AE' => 5.33, 'AF' => 5.33,
            'AG' => 5.33, 'AH' => 12, 'AI' => 54.33
        ];
    }
}