<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class ResponseTimeExport implements FromCollection, WithEvents, WithColumnWidths, WithTitle
{
    protected array $data;
    protected string $sheetName;

    protected int $lastTableRow = 13;

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
        $sheet->mergeCells('B1:D3');
        // buat objek image
        $leftLogo = new Drawing();
        $leftLogo->setName('Image');
        $leftLogo->setDescription('Image B1:D3');
        $leftLogo->setPath(public_path('ids.png'));

        // ukuran image (pixel)
        $leftLogo->setWidth(170);

        // anchor ke cell hasil merge
        $leftLogo->setCoordinates('B1');

        // optional: center-kan di merge cell
        $leftLogo->setOffsetX(0);
        $leftLogo->setOffsetY(0);

        $leftLogo->setWorksheet($sheet);

        // buat objek image
        $rightLogo = new Drawing();
        $rightLogo->setName('Image');
        $rightLogo->setDescription('Image K1:K3');
        $rightLogo->setPath(public_path('pertamina-patra-logistik.png'));

        // ukuran image (pixel)
        $rightLogo->setWidth(165);

        // anchor ke cell hasil merge
        $rightLogo->setCoordinates('K1');

        // optional: center-kan di merge cell
        $rightLogo->setOffsetX(60);
        $rightLogo->setOffsetY(0);

        $rightLogo->setWorksheet($sheet);

        $sheet->setCellValue('B4', 'LAMPIRAN 6');
        $sheet->setCellValue('B5', 'REKAPITULASI TINDAKLANJUT RESPONSE TIME KELUHAN USER/STAKEHOLDER');
        $sheet->setCellValue('B6', 'PERIODE : SEPTEMBER 2025');
        $sheet->setCellValue('B7', 'LOKASI    : IT Jakarta');

        $sheet->getStyle('B4:B7')->getFont()->setBold(true)->setSize(12);
        $sheet->getStyle('B4')->getFont()->setSize(16);

        $sheet->getRowDimension(4)->setRowHeight(21);
     }

     /* =====================================================
     | FOOTER
     ===================================================== */
     private function renderFooter($sheet)
     {
        $sheet->mergeCells("C15:E15");
        $sheet->getRowDimension(15)->setRowHeight(156);

        $richText = new RichText();
        $boldText = $richText->createTextRun('Tata Cara Perhitungan KPI:');
        $boldText->getFont()->setBold(true)->setUnderline(true);
        $richText->createText("\n");
        
        $run = $richText->createTextRun('Kriteria 1 - penyelesaian < target waktu penyelesaian = 110%');
        $run->getFont()->setItalic(true);

        $richText->createText("\n");

        $run = $richText->createTextRun('Kriteria 2 - lebih 2 hari dari target = 95%');
        $run->getFont()->setItalic(true);

        $richText->createText("\n");

        $run = $richText->createTextRun('Kriteria 3 - lebih 4 hari dari target = 90%');
        $run->getFont()->setItalic(true);

        $richText->createText("\n");

        $run = $richText->createTextRun('Kriteria 4 - lebih 6 hari dari target = 85%');
        $run->getFont()->setItalic(true);

        $richText->createText("\n");

        $run = $richText->createTextRun('Dst, setiap kelipatan 2 hari keterlambatan');
        $run->getFont()->setItalic(true);

        $richText->createText("\n");

        $run = $richText->createTextRun('dari target akan mengurangi 5%');
        $run->getFont()->setItalic(true);

        $richText->createText("\n\n");

        $run = $richText->createTextRun('Selama belum melewati due date maka perhitungan parameter ini diasumsikan 100%');
        $run->getFont()->setItalic(true);

        $sheet->setCellValue("C15", $richText);
        $sheet->getStyle("C15")->getFont()->setSize(11)->setItalic(true);
        $sheet->getStyle("C15")->getAlignment()->setVertical(Alignment::VERTICAL_TOP)->setWrapText(true);

        // Tanda Tangan
        $signatureRow1 = 18;
        $signatureRow2 = $signatureRow1 + 1;
        $signatureRow3 = $signatureRow2 + 1;
        $signatureRow4 = $signatureRow3 + 7;

        $sheet->setCellValue("B{$signatureRow1}", 'Disiapkan Oleh,');
        $sheet->setCellValue("B{$signatureRow2}", 'PT Indi Daya Sistem');
        $sheet->setCellValue("B{$signatureRow3}", 'Supervisor Services');
        $sheet->setCellValue("B{$signatureRow4}", 'Dimas Nafidin');

        $sheet->setCellValue("H{$signatureRow1}", 'Disetujui Oleh,');
        $sheet->setCellValue("H{$signatureRow2}", 'PT  Patra Logistik');
        $sheet->setCellValue("H{$signatureRow3}", 'Area Manager Jawa Bagian Barat');
        $sheet->setCellValue("H{$signatureRow4}", 'Bayu Riyadi');

        $sheet->getStyle("B{$signatureRow2}:B{$signatureRow4}")->getFont()->setBold(true);
        $sheet->getStyle("H{$signatureRow2}:H{$signatureRow4}")->getFont()->setBold(true);

        $sheet->getStyle("B{$signatureRow1}:H{$signatureRow4}")->getFont()->setSize(14);
     }

     /* =====================================================
     | RENDER TABLE
     ===================================================== */
    private function renderTable($sheet, array $config): void
    {
        $startColumnIndex = Coordinate::columnIndexFromString('B');

        $column1 = Coordinate::stringFromColumnIndex($startColumnIndex);        // No
        $column2 = Coordinate::stringFromColumnIndex($startColumnIndex + 1);    // Tanggal Keluhan
        $column3 = Coordinate::stringFromColumnIndex($startColumnIndex + 2);    // Nomor Keluhan
        $column4 = Coordinate::stringFromColumnIndex($startColumnIndex + 3);    // User
        $column5 = Coordinate::stringFromColumnIndex($startColumnIndex + 4);    // Rekomendasi
        $column6 = Coordinate::stringFromColumnIndex($startColumnIndex + 5);    // Jumlah Temuan
        $column7 = Coordinate::stringFromColumnIndex($startColumnIndex + 6);    // Tanggal Temuan
        $column8 = Coordinate::stringFromColumnIndex($startColumnIndex + 7);    // Realisasi Closing
        $column9 = Coordinate::stringFromColumnIndex($startColumnIndex + 8);    // Jumlah Hari Tindak Lanjut
        $column10 = Coordinate::stringFromColumnIndex($startColumnIndex + 9);   // Tindak Lanjut
        $column11 = Coordinate::stringFromColumnIndex($startColumnIndex + 10);  // Nilai

        $headerRow1 = 9;
        $headerRow2 = $headerRow1 + 1;
        
        $bodyStartRow = $headerRow2 + 1;

        // HEADER
        $sheet->mergeCells("{$column1}{$headerRow1}:{$column1}{$headerRow2}")->setCellValue("{$column1}{$headerRow1}", "No");
        $sheet->mergeCells("{$column2}{$headerRow1}:{$column2}{$headerRow2}")->setCellValue("{$column2}{$headerRow1}", "Tanggal Keluhan");
        $sheet->mergeCells("{$column3}{$headerRow1}:{$column3}{$headerRow2}")->setCellValue("{$column3}{$headerRow1}", "Nomor Keluhan");
        $sheet->mergeCells("{$column4}{$headerRow1}:{$column4}{$headerRow2}")->setCellValue("{$column4}{$headerRow1}", "User");
        $sheet->mergeCells("{$column5}{$headerRow1}:{$column5}{$headerRow2}")->setCellValue("{$column5}{$headerRow1}", "Rekomendasi");
        $sheet->mergeCells("{$column6}{$headerRow1}:{$column6}{$headerRow2}")->setCellValue("{$column6}{$headerRow1}", "Jumlah Temuan");

        $sheet->mergeCells("{$column7}{$headerRow1}:{$column8}{$headerRow1}")->setCellValue("{$column7}{$headerRow1}", "Due Date");
        $sheet->setCellValue("{$column7}{$headerRow2}", "Tanggal Temuan");
        $sheet->setCellValue("{$column8}{$headerRow2}", "Realisasi Closing");

        $sheet->mergeCells("{$column9}{$headerRow1}:{$column9}{$headerRow2}")->setCellValue("{$column9}{$headerRow1}", "Jumlah Hari Tindak Lanjut");
        $sheet->mergeCells("{$column10}{$headerRow1}:{$column10}{$headerRow2}")->setCellValue("{$column10}{$headerRow1}", "Tindak Lanjut");
        $sheet->mergeCells("{$column11}{$headerRow1}:{$column11}{$headerRow2}")->setCellValue("{$column11}{$headerRow1}", "Nilai");

        $this->applyHeaderStyle($sheet, "{$column1}{$headerRow1}:{$column11}{$headerRow2}");
        $sheet->getRowDimension($headerRow1)->setRowHeight(27);


        // BODY
        $sheet->fromArray([
            [1, "-", "-", "-", "-", "-", "-", "-", "-", null, 110]
        ], null, "{$column1}{$bodyStartRow}");

        $this->applyBodyStyle($sheet, "{$column1}{$bodyStartRow}:{$column11}{$bodyStartRow}", Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle("{$column5}{$bodyStartRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        $sheet->getRowDimension($bodyStartRow)->setRowHeight(77);

        $this->applyBackgroundColor($sheet, "{$column9}{$bodyStartRow}", "c6efcd");

        // FOOTER
        $footerRow1 = 12;
        $footerRow2 = $footerRow1 + 1;

        $sheet->mergeCells("{$column7}{$footerRow1}:{$column10}{$footerRow1}")->setCellValue("{$column7}{$footerRow1}", "TOTAL TEMUAN");
        $sheet->mergeCells("{$column7}{$footerRow2}:{$column10}{$footerRow2}")->setCellValue("{$column7}{$footerRow2}", "TOTAL NILAI KPI");

        $sheet->setCellValue("{$column11}{$footerRow2}", 110);

        $this->applyBodyStyle($sheet, "{$column7}{$footerRow1}:{$column11}{$footerRow2}");

        $sheet->getStyle("{$column11}{$footerRow1}:{$column11}{$footerRow2}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER);

        $sheet->getStyle("{$column7}{$footerRow1}:{$column7}{$footerRow2}")->getFont()->setBold(true);


        $sheet->getStyle("{$column7}{$footerRow1}:{$column11}{$footerRow2}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_MEDIUM);
        $sheet->getStyle("{$column1}{$bodyStartRow}:{$column11}{$bodyStartRow}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_MEDIUM);
        $sheet->getStyle("{$column1}{$headerRow1}:{$column11}{$headerRow2}")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_MEDIUM);

        $sheet->getStyle("{$column11}{$bodyStartRow}:{$column11}{$footerRow2}")->getNumberFormat()->setFormatCode('#,##0"%"');
        // $startColIndex = Coordinate::columnIndexFromString('E');
        // $endColIndex   = $startColIndex + $totalDate - 1;

        // $startColDate = Coordinate::stringFromColumnIndex($startColIndex);
        // $endColDate   = Coordinate::stringFromColumnIndex($endColIndex);

        // $columnTotal = Coordinate::stringFromColumnIndex($endColIndex + 1);

        // $startCol = $config['startCol'];
        // $endCol   = $config['endCol'];
        // $startRow = $config['startRow'];
        // $data     = $config['data'] ?? [];
        // $minRows  = count($data) > 10 ? count($data) : 10; // Minimum 10 baris body
        
        // $headerRow1 = $startRow;
        // $headerRow2 = $headerRow1 + 1;
        // $bodyStart  = $headerRow2 + 1;
        // $bodyEnd    = $bodyStart + $minRows - 1;
        // $footerRow1 = $bodyEnd + 1;
        // $footerRow2 = $footerRow1 + 1;

        // // ================= HEADER =================
        // $column1 = $startCol;
        // $column2 = chr(ord($column1) + 1);
        // $column3 = chr(ord($column2) + 1);
        // $column4 = chr(ord($column3) + 1);
        // $column5 = chr(ord($column4) + 1);
        // $column6 = chr(ord($column5) + 1);

        // // Custom Text Bold
        // $richText = new RichText();
        // $richText->createText('Selisih Waktu');
        // $richText->createText("\n");
        // $normalText = $richText->createTextRun('(Day:Jam:Menit:Detik)');
        // $normalText->getFont()->setBold(false);

        // $column = [
        //     ['value' => $column1, 'title' => 'No.'],
        //     ['value' => $column2, 'title' => 'Tanggal'],
        //     ['value' => $column3, 'title' => 'Target Pelaporan'],
        //     ['value' => $column4, 'title' => 'Realisasi Laporan'],
        //     ['value' => $column5, 'title' => $richText],
        //     ['value' => $column6, 'title' => 'Persentase'],
        // ];

        // foreach ($column as $key => $col) {
        //     $sheet->mergeCells("{$col['value']}{$headerRow1}:{$col['value']}{$headerRow2}");
        //     $sheet->setCellValue("{$col['value']}{$headerRow1}", $col['title']);
        // }

        // $this->applyHeaderStyle($sheet, "{$column1}{$headerRow1}:{$column6}{$headerRow2}");


        // // ================= BODY =================
        // // Render data yang ada
        // for ($i = 0; $i < $minRows; $i++) {
        //     $currentRow = $bodyStart + $i;
            
        //     if (isset($data[$i])) {
        //         // Ada data, render data asli
        //         $item = $data[$i];
        //         $sheet->setCellValue("{$column1}{$currentRow}", $i+1);
        //         $sheet->setCellValue("{$column2}{$currentRow}", $item['tanggal']);
        //         $sheet->setCellValue("{$column3}{$currentRow}", $item['target_pelaporan']);
        //         $sheet->setCellValue("{$column4}{$currentRow}", $item['realisasi_laporan']);
        //         $sheet->setCellValue("{$column5}{$currentRow}", $item['selisih_waktu']);
        //         $sheet->setCellValue("{$column6}{$currentRow}", $item['persentase'].'%');

        //         $sheet->getStyle("{$column6}{$currentRow}")->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_PERCENTAGE);

        //         $this->applyHorizontalAlignment($sheet, "{$column1}{$currentRow}:{$column6}{$currentRow}", Alignment::HORIZONTAL_RIGHT);
        //         $sheet->getStyle("{$column1}{$currentRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        //         $sheet->getStyle("{$column6}{$currentRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        //     } else {
        //         // Tidak ada data, render baris kosong
        //         $sheet->setCellValue("{$column1}{$currentRow}", '');
        //         $sheet->setCellValue("{$column2}{$currentRow}", '');
        //         $sheet->setCellValue("{$column3}{$currentRow}", '');
        //         $sheet->setCellValue("{$column4}{$currentRow}", '');
        //         $sheet->setCellValue("{$column5}{$currentRow}", '');
        //         $sheet->setCellValue("{$column6}{$currentRow}", '');
        //     }
        // }


        // // Catatan
        // $row1 = $startRow;
        // $row2 = $row1 + 1;
        // $row3 = $row2 + 1;

        // $sheet->setCellValue("I{$row1}", 'Catatan:');
        // $sheet->setCellValue("I{$row2}", '1 Jam = ');
        // $sheet->setCellValue("I{$row3}", '1 Jam 1 Dtk =');

        // $sheet->setCellValue("J{$row2}", 0.0416666666715173);
        // $sheet->setCellValue("J{$row3}", 0.041678240741021);

        // $sheet->getStyle("I{$row1}:J{$row3}")
        // ->getBorders()
        // ->getOutline()
        // ->setBorderStyle(Border::BORDER_MEDIUM);

        // $sheet->getStyle("I{$row1}")->getFont()->setBold(true);


        // // ================= FOOTER =================
        // $sheet->setCellValue("{$column1}{$footerRow1}", '');
        
        // $sheet->mergeCells("{$column1}{$footerRow2}:{$column5}{$footerRow2}");
        // $sheet->setCellValue("{$column1}{$footerRow2}", 'Nilai KPI');
        
        // $sheet->getStyle("{$column1}{$footerRow2}:{$column6}{$footerRow2}")->applyFromArray([
        //     'font' => ['bold' => true, 'color' => ['rgb' => '000000']],
        //     'alignment' => [
        //         'horizontal' => Alignment::HORIZONTAL_CENTER,
        //         'vertical' => Alignment::VERTICAL_CENTER,
        //     ],
        //     'fill' => [
        //         'fillType' => Fill::FILL_SOLID,
        //         'startColor' => ['rgb' => 'ffbf00'],
        //     ],
        //     'borders' => [
        //         'allBorders' => ['borderStyle' => Border::BORDER_THIN],
        //     ],
        // ]);

        // $this->lastTableRow = $footerRow2;

        // $sheet->setCellValue("{$column6}{$footerRow2}", '110%');

        // // Apply border ke semua 10 baris
        // $this->applyBorder($sheet, "{$startCol}{$bodyStart}:{$endCol}{$footerRow2}");
    }

    /* =====================================================
     | STYLES
     ===================================================== */
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
            'font' => ['bold' => true, 'color' => ['rgb' => '000000'], 'size' => 11],
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
        $this->applyVerticalAlignment($sheet, $range);
    }

    private function applyHorizontalAlignment($sheet, string $range, string $horizontalAlignment = Alignment::HORIZONTAL_LEFT): void
    {
        $sheet->getStyle($range)->getAlignment()->setHorizontal($horizontalAlignment);
    }

    private function applyVerticalAlignment($sheet, string $range, string $verticalAlignment = Alignment::VERTICAL_CENTER): void
    {
        $sheet->getStyle($range)->getAlignment()->setVertical($verticalAlignment);
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
            'A' => 2.67, 'B' => 2.67, 'C' => 10, 'D' => 14.33, 'E' => 20.50,
            'F' => 45.50, 'G' => 8.83, 'H' => 14.33, 'I' => 14.33, 'J' => 10.50,
            'K' => 29.83, 'L' => 6
        ];

        return array_map(fn($v) => $v + 0.83, $data);
    }
}