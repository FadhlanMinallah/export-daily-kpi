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
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class SummaryExport implements FromCollection, WithEvents, WithColumnWidths, WithTitle
{
    protected array $data;
    protected string $sheetName;

    public function __construct(array $data, string $sheetName = 'Summary')
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

                $this->renderTable($sheet);

                // $sheet->setShowGridlines(false);
            }
        ];
    }

     /* =====================================================
     | RENDER TABLE
     ===================================================== */
    private function renderTable($sheet): void
    {
        $richText = new RichText();

        $sheet->setCellValue('B1', 'III. Contoh Perhitungan Denda KPI');

        $sheet->setCellValue('A2', 'a. Pengurangan pembayaran bulanan diberikan jika KONTRAKTOR mendapatkan nilai di bawah target KPI yang ditetapkan sesuai ketentuan dimana nilai pengurangan dihitung dari presentase selisih dari jumlah nilai ketidaktercapaian KPI.');
        $sheet->setCellValue('A3', 'b. Nilai Key Performance Indicator (KPI) dari IT/FT/Lokasi setiap bulan sesuai dengan kesepakatan para pihak.');
        // $sheet->setCellValue('A4', 'c. Maksimal Denda KPI sebesar 5% (lima persen) dari Nilai Kontrak.');

        // Custom Text Bold
        $richText->createText(
            'c. Maksimal Denda KPI sebesar 5% '
        );
        $boldText = $richText->createTextRun('(lima persen)');
        $boldText->getFont()->setBold(true);
        $richText->createText(
            ' dari Nilai Kontrak.'
        );
        $sheet->setCellValue('A4', $richText);

        $sheet->setCellValue('A7', 'II. Rumusan Denda KPI');
        $sheet->setCellValue('A8', 'Denda KPI = 5% x (100% - Pencapaian KPI) x Tagihan Bulanan');
        $sheet->getStyle('A8')->getFont()->setItalic(true);

        $sheet->setCellValue('B2', 'Rumus Denda KPI = 5% x (100% - Pencapaian KPI) x Tagihan Bulanan');
        $sheet->getStyle('B2')->getFont()->setItalic(true);

        $sheet->fromArray([
            ['Simulasi:', null, null],
            ['Jumlah GPS', '248', 'Unit'],
            ['Jumlah Dashcam', '248', 'Unit'],
            ['Tarif GPS', 350000, 'Rp/Unit/Bulan'],
            ['Tarif Dashcam', 566000, 'Rp/Unit/Bulan'],
            ['Pembayaran GPS & Dashcam', 227168000, 'Rp/Unit/Bulan'],
            ['Nilai KPI', 95.5, '%'],
            ['Denda KPI', 513749, 'Rp'],
            [null, null, null],
            ['Nilai denda KPI+Perangkat tidak Ops Dashcam+GPS', 2455749, 513749, 513749, 1942000],
            [null, null, null],
            ['Nilai Tagihan=', 'Total Tagihan GPS dan Dashcam-Nilai Denda'],
            ['', 224712251],
            [null, null],
            ['Denda Ops', 1942000],
            ['Denda Ins', null]
        ], null, 'B4');

        $sheet->getStyle('A2:A8')->getAlignment()->setWrapText(false);
        
        $sheet->getStyle('A7')->getFont()->setBold(true);
        $sheet->getStyle('B1')->getFont()->setBold(true);
        $sheet->getStyle('B4')->getFont()->setBold(true);
        $sheet->getStyle('B9:D9')->getFont()->setBold(true);
        $sheet->getStyle('C11')->getFont()->setBold(true);
        $sheet->getStyle('C13')->getFont()->setBold(true);
        $sheet->getStyle('C16')->getFont()->setBold(true);
        $sheet->getStyle('F13')->getFont()->setBold(true);

        $sheet->getStyle('B11:D11')->applyFromArray([
            'font' => ['color' => ['rgb' => 'c00000']],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'f2f2f2'],
            ],
        ]);

        $sheet->getStyle('B13')->getAlignment()->setWrapText(false);

        $sheet->getStyle('C7:C9')->getNumberFormat()->setFormatCode('#,##0');
        $sheet->getStyle('C11')->getNumberFormat()->setFormatCode('#,##0');
        $sheet->getStyle('C13')->getNumberFormat()->setFormatCode('#,##0');
        $sheet->getStyle('C16')->getNumberFormat()->setFormatCode('#,##0');

        // $sheet->getStyle('C10')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_PERCENTAGE);

        $sheet->getStyle('C10')->applyFromArray([
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'ffff00'],
            ],
        ]);

        $sheet->getStyle('D13:F13')->getNumberFormat()->setFormatCode('"Rp"  #,##0');
        $sheet->getStyle('C18:C19')->getNumberFormat()->setFormatCode('"Rp"  #,##0');
    }

    /* =====================================================
     | STYLES
     ===================================================== */
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
            ->getFont()->setName('Arial');
    }

    /* =====================================================
     | COLUMN WIDTHS
     ===================================================== */
    public function columnWidths(): array
    {
        return [
            'A' => 97.50, 'B' => 27, 'C' => 12.83, 'D' => 12.50,
            'E' => 14.67, 'F' => 13, 'G' => 9.83
        ];
    }
}