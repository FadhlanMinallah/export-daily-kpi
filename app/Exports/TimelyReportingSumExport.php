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

class TimelyReportingSumExport implements FromCollection, WithEvents, WithColumnWidths, WithTitle
{
    protected array $data;
    protected string $sheetName;

    protected int $lastTableRow;

    public function __construct(array $data, string $sheetName = 'Timely_Reporting_Sum')
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

                $this->renderHeader($sheet);

                $this->renderTable($sheet, [
                    'startCol'  => 'A',
                    'endCol'    => 'F',
                    'startRow'  => 9,
                    'data'      => $this->data['item'] ?? []
                ]);

                $this->renderFooter($sheet);

                // $sheet->setShowGridlines(false);
            }
        ];
    }

     /* =====================================================
     | HEADER
     ===================================================== */
     private function renderHeader($sheet)
     {
        $sheet->mergeCells('A1:C3');
        // buat objek image
        $leftLogo = new Drawing();
        $leftLogo->setName('Image');
        $leftLogo->setDescription('Image A1:C3');
        $leftLogo->setPath(public_path('ids.png'));

        // ukuran image (pixel)
        $leftLogo->setWidth(170);

        // anchor ke cell hasil merge
        $leftLogo->setCoordinates('A1');

        // optional: center-kan di merge cell
        $leftLogo->setOffsetX(0);
        $leftLogo->setOffsetY(0);

        $leftLogo->setWorksheet($sheet);

        // buat objek image
        $rightLogo = new Drawing();
        $rightLogo->setName('Image');
        $rightLogo->setDescription('Image E1:F3');
        $rightLogo->setPath(public_path('pertamina-patra-logistik.png'));

        // ukuran image (pixel)
        $rightLogo->setWidth(165);

        // anchor ke cell hasil merge
        $rightLogo->setCoordinates('E1');

        // optional: center-kan di merge cell
        $rightLogo->setOffsetX(60);
        $rightLogo->setOffsetY(0);

        $rightLogo->setWorksheet($sheet);

        $sheet->setCellValue('A4', 'LAMPIRAN 3C');
        $sheet->setCellValue('A5', 'REKAPITULASI LAPORAN BULANAN OPERASIONAL GPS & DASHCAM');
        $sheet->setCellValue('A6', 'PERIODE : SEPTEMBER 2025');
        $sheet->setCellValue('A7', 'LOKASI    : IT Jakarta');

        $sheet->getStyle('A4:A7')->getFont()->setBold(true)->setSize(9);
        $sheet->getStyle('A4')->getFont()->setSize(10);

     }

     /* =====================================================
     | FOOTER
     ===================================================== */
     private function renderFooter($sheet)
     {
        // $sheet->setCellValue('A23', $this->lastTableRow);

        $startRow = $this->lastTableRow + 2;
        $row2 = $startRow + 1;

        $sheet->setCellValue("B{$startRow}", 'Tata Cara Perhitungan KPI:');

        $sheet->mergeCells("B{$startRow}:C{$startRow}");
        $sheet->mergeCells("B{$row2}:C{$row2}");
        $sheet->getRowDimension($row2)->setRowHeight(52);

        $richText = new RichText();
        $richText->createText('Kriteria :');
        $richText->createText("\n");
        $richText->createText('Kriteria 1 - Pukul 07.00 LT Day +1 = 110%');
        $richText->createText("\n");
        $richText->createText('Kriteria 2 - Pukul 08.00 LT Day +1 = 100 %');
        $richText->createText("\n");
        $richText->createText('Kriteria 3 - Lebih dari Pukul 09.00 LT Day +1 = 80 %');


        $sheet->setCellValue("B{$row2}", $richText);
        $sheet->getStyle("B{$startRow}")->getFont()->setBold(true);
        $sheet->getStyle("B{$startRow}:B{$row2}")->getFont()->setSize(8);
        $sheet->getStyle("B{$row2}")->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

        $sheet->getStyle("B{$row2}:C{$row2}")->getAlignment()->setWrapText(true);

        // Tanda Tangan
        $signatureRow1 = $row2 + 3;
        $signatureRow2 = $signatureRow1 + 1;
        $signatureRow3 = $signatureRow2 + 1;
        $signatureRow4 = $signatureRow3 + 7;

        $sheet->setCellValue("B{$signatureRow1}", 'Disiapkan Oleh,');
        $sheet->setCellValue("B{$signatureRow2}", 'PT Indi Daya Sistem');
        $sheet->setCellValue("B{$signatureRow3}", 'Supervisor Services');
        $sheet->setCellValue("B{$signatureRow4}", 'Dimas Nafidin');

        $sheet->setCellValue("E{$signatureRow1}", 'Disetujui Oleh,');
        $sheet->setCellValue("E{$signatureRow2}", 'PT  Patra Logistik');
        $sheet->setCellValue("E{$signatureRow3}", 'Area Manager Jawa Bagian Barat');
        $sheet->setCellValue("E{$signatureRow4}", 'Bayu Riyadi');

        $sheet->getStyle("B{$signatureRow2}:B{$signatureRow4}")->getFont()->setBold(true);
        $sheet->getStyle("E{$signatureRow2}:E{$signatureRow4}")->getFont()->setBold(true);

        $sheet->getStyle("B{$signatureRow1}:E{$signatureRow4}")->getFont()->setSize(14);
     }

     /* =====================================================
     | RENDER TABLE
     ===================================================== */
    private function renderTable($sheet, array $config): void
    {
        $startCol = $config['startCol'];
        $endCol   = $config['endCol'];
        $startRow = $config['startRow'];
        $data     = $config['data'] ?? [];
        $minRows  = count($data) > 10 ? count($data) : 10; // Minimum 10 baris body
        
        $headerRow1 = $startRow;
        $headerRow2 = $headerRow1 + 1;
        $bodyStart  = $headerRow2 + 1;
        $bodyEnd    = $bodyStart + $minRows - 1;
        $footerRow1 = $bodyEnd + 1;
        $footerRow2 = $footerRow1 + 1;

        // ================= HEADER =================
        $column1 = $startCol;
        $column2 = chr(ord($column1) + 1);
        $column3 = chr(ord($column2) + 1);
        $column4 = chr(ord($column3) + 1);
        $column5 = chr(ord($column4) + 1);
        $column6 = chr(ord($column5) + 1);

        // Custom Text Bold
        $richText = new RichText();
        $richText->createText('Selisih Waktu');
        $richText->createText("\n");
        $normalText = $richText->createTextRun('(Day:Jam:Menit:Detik)');
        $normalText->getFont()->setBold(false);

        $column = [
            ['value' => $column1, 'title' => 'No.'],
            ['value' => $column2, 'title' => 'Tanggal'],
            ['value' => $column3, 'title' => 'Target Pelaporan'],
            ['value' => $column4, 'title' => 'Realisasi Laporan'],
            ['value' => $column5, 'title' => $richText],
            ['value' => $column6, 'title' => 'Persentase'],
        ];

        foreach ($column as $key => $col) {
            $sheet->mergeCells("{$col['value']}{$headerRow1}:{$col['value']}{$headerRow2}");
            $sheet->setCellValue("{$col['value']}{$headerRow1}", $col['title']);
        }

        $this->applyHeaderStyle($sheet, "{$column1}{$headerRow1}:{$column6}{$headerRow2}");


        // ================= BODY =================
        // Render data yang ada
        for ($i = 0; $i < $minRows; $i++) {
            $currentRow = $bodyStart + $i;
            
            if (isset($data[$i])) {
                // Ada data, render data asli
                $item = $data[$i];
                $sheet->setCellValue("{$column1}{$currentRow}", $i+1);
                $sheet->setCellValue("{$column2}{$currentRow}", $item['tanggal']);
                $sheet->setCellValue("{$column3}{$currentRow}", $item['target_pelaporan']);
                $sheet->setCellValue("{$column4}{$currentRow}", $item['realisasi_laporan']);
                $sheet->setCellValue("{$column5}{$currentRow}", $item['selisih_waktu']);
                $sheet->setCellValue("{$column6}{$currentRow}", $item['persentase'].'%');

                $sheet->getStyle("{$column6}{$currentRow}")->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_PERCENTAGE);

                $this->applyHorizontalAlignment($sheet, "{$column1}{$currentRow}:{$column6}{$currentRow}", Alignment::HORIZONTAL_RIGHT);
                $sheet->getStyle("{$column1}{$currentRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $sheet->getStyle("{$column6}{$currentRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            } else {
                // Tidak ada data, render baris kosong
                $sheet->setCellValue("{$column1}{$currentRow}", '');
                $sheet->setCellValue("{$column2}{$currentRow}", '');
                $sheet->setCellValue("{$column3}{$currentRow}", '');
                $sheet->setCellValue("{$column4}{$currentRow}", '');
                $sheet->setCellValue("{$column5}{$currentRow}", '');
                $sheet->setCellValue("{$column6}{$currentRow}", '');
            }
        }


        // Catatan
        $row1 = $startRow;
        $row2 = $row1 + 1;
        $row3 = $row2 + 1;

        $sheet->setCellValue("I{$row1}", 'Catatan:');
        $sheet->setCellValue("I{$row2}", '1 Jam = ');
        $sheet->setCellValue("I{$row3}", '1 Jam 1 Dtk =');

        $sheet->setCellValue("J{$row2}", 0.0416666666715173);
        $sheet->setCellValue("J{$row3}", 0.041678240741021);

        $sheet->getStyle("I{$row1}:J{$row3}")
        ->getBorders()
        ->getOutline()
        ->setBorderStyle(Border::BORDER_MEDIUM);

        $sheet->getStyle("I{$row1}")->getFont()->setBold(true);


        // ================= FOOTER =================
        $sheet->setCellValue("{$column1}{$footerRow1}", '');
        
        $sheet->mergeCells("{$column1}{$footerRow2}:{$column5}{$footerRow2}");
        $sheet->setCellValue("{$column1}{$footerRow2}", 'Nilai KPI');
        
        $sheet->getStyle("{$column1}{$footerRow2}:{$column6}{$footerRow2}")->applyFromArray([
            'font' => ['bold' => true, 'color' => ['rgb' => '000000']],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'ffbf00'],
            ],
            'borders' => [
                'allBorders' => ['borderStyle' => Border::BORDER_THIN],
            ],
        ]);

        $this->lastTableRow = $footerRow2;

        $sheet->setCellValue("{$column6}{$footerRow2}", '110%');

        // Apply border ke semua 10 baris
        $this->applyBorder($sheet, "{$startCol}{$bodyStart}:{$endCol}{$footerRow2}");
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
                'startColor' => ['rgb' => '8db3e2'],
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
        $data = [
            'A' => 2.67, 'B' => 9.50, 'C' => 18, 'D' => 18, 'E' => 17.80,
            'F' => 9.33, 'G' => 9.83, 'H' => 14.33, 'I' => 9.50, 'J' => 10.33
        ];

        return array_map(fn($v) => $v + 0.83, $data);
    }
}