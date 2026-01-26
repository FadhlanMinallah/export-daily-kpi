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
                $drawing->setPath(public_path('pertamina-patra-logistik.png'));

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
        // ===== HEADER =====
        $sheet->mergeCells('A6:A7')->setCellValue('A6', 'No');
        $sheet->mergeCells('B6:B7')->setCellValue('B6', 'Item Check');
        $sheet->mergeCells('C6:AG6')->setCellValue('C6', 'Jan-26');

        // tanggal 1–31
        $col = 'C';
        for ($i = 1; $i <= 31; $i++) {
            $sheet->setCellValue($col.'7', $i);
            $col++;
        }

        $sheet->mergeCells('AH6:AH7')->setCellValue('AH6', "Rata-rata\nBulanan");
        $sheet->mergeCells('AI6:AI7')->setCellValue('AI6', 'Keterangan');

        $sheet->getStyle('AH6')->getAlignment()->setWrapText(true);
        $sheet->getStyle('A6:AI7')->getFont()->setBold(true)->setSize(10);
        $this->applyHeaderStyle($sheet, 'A6:AI7');

        // ===== BODY =====
        $rowStart = 8;
        $no = 1;

        foreach ($this->data['items'] as $item) {

            // No
            // $sheet->setCellValue("A{$rowStart}", $no);

            // Item Check
            $sheet->setCellValue("B{$rowStart}", $item['item']);

            // isi tanggal
            $col = 'C';
            for ($d = 1; $d <= 31; $d++) {

                $value = $item['values'][$d] ?? null;

                if ($value !== null) {
                    $sheet->setCellValue($col.$rowStart, $value);

                    if ($item['is_percent']) {
                        $sheet->getStyle($col.$rowStart)
                            ->getNumberFormat()
                            ->setFormatCode(NumberFormat::FORMAT_PERCENTAGE);
                    }
                }

                $col++;
            }

            // Rata-rata bulanan (FORMULA)
            $sheet->setCellValue(
                "AH{$rowStart}",
                "=IFERROR(AVERAGE(C{$rowStart}:AG{$rowStart}),\"#DIV/0!\")"
            );

            // Keterangan
            $sheet->setCellValue("AI{$rowStart}", $item['keterangan']);

            // style body
            $this->applyBorder($sheet, "A{$rowStart}:AI{$rowStart}");
            $sheet->getStyle("A{$rowStart}:AH{$rowStart}")
                ->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER);

            $sheet->getStyle("B{$rowStart}:AI{$rowStart}")
                ->getAlignment()
                ->setVertical(Alignment::VERTICAL_CENTER);

            if ($item['is_percent']) {
                $sheet->getStyle("B{$rowStart}:AI{$rowStart}")->applyFromArray([
                    'font' => ['bold' => true, 'color' => ['rgb' => '000000']],
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => ['rgb' => '71a8e0'],
                    ],
                ]);
            }

            $rowStart++;
            $no++;
        }
        $this->applyHorizontalAlignment($sheet, 'B8:B21', Alignment::HORIZONTAL_LEFT);
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
            'A' => 5.50, 'B' => 43.33, 'C' => 5.80, 'D' => 5.80,
            'E' => 5.80, 'F' => 5.80, 'G' => 5.80, 'H' => 5.80,
            'I' => 5.80, 'J' => 5.80, 'K' => 5.80, 'L' => 5.80,
            'M' => 5.80, 'N' => 5.80, 'O' => 5.80, 'P' => 5.80, 'Q' => 5.80,
            'R' => 5.80, 'S' => 5.80, 'T' => 5.80, 'U' => 5.80, 'V' => 5.80,
            'W' => 5.80, 'X' => 5.80, 'Y' => 5.80, 'Z' => 5.80, 'AA' => 5.80,
            'AB' => 5.80, 'AC' => 5.80, 'AD' => 5.80, 'AE' => 5.80, 'AF' => 5.80,
            'AG' => 5.80, 'AH' => 12, 'AI' => 54.33
        ];
    }
}