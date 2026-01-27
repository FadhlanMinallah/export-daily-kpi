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

class BappExport implements FromCollection, WithEvents, WithColumnWidths, WithTitle
{
    protected array $data;
    protected string $sheetName;

    protected int $lastTableRow;

    public function __construct(array $data, string $sheetName = 'BAPP')
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
                    'startCol'  => 'B',
                    'endCol'    => 'N',
                    'startRow'  => 13,
                    'data'      => $this->data ?? []
                ]);

                // $this->renderFooter($sheet);

                // $sheet->setShowGridlines(false);
            }
        ];
    }

     /* =====================================================
     | HEADER
     ===================================================== */
     private function renderHeader($sheet)
     {
        $sheet->mergeCells('B4:N9');
        // buat objek image
        $leftLogo = new Drawing();
        $leftLogo->setName('Image');
        $leftLogo->setDescription('Image B4:C6');
        $leftLogo->setPath(public_path('ids.png'));

        // ukuran image (pixel)
        $leftLogo->setWidth(180);

        // anchor ke cell hasil merge
        $leftLogo->setCoordinates('B4');

        // optional: center-kan di merge cell
        $leftLogo->setOffsetX(20);
        $leftLogo->setOffsetY(20);

        $leftLogo->setWorksheet($sheet);

        // buat objek image
        $rightLogo = new Drawing();
        $rightLogo->setName('Image');
        $rightLogo->setDescription('Image L4:N6');
        $rightLogo->setPath(public_path('pertamina-patra-logistik.png'));

        // ukuran image (pixel)
        $rightLogo->setWidth(180);

        // anchor ke cell hasil merge
        $rightLogo->setCoordinates('L4');

        // optional: center-kan di merge cell
        $rightLogo->setOffsetX(55);
        $rightLogo->setOffsetY(15);

        $rightLogo->setWorksheet($sheet);

        $richText = new RichText();
        $richText->createText('BERITA ACARA');
        $richText->createText("\n");
        $richText->createText('PENYELESAIAN PEKERJAAN');
        $richText->createText("\n");
        $richText->createText('JASA SEWA GPS DAN DASHCAM');
        $richText->createText("\n");
        $richText->createText('INTEGRATED TERMINAL JAKARTA PERIODE SEPTEMBER');

        $sheet->setCellValue('B4', $richText);
        $sheet->getStyle("B4")->getFont()->setBold(true)->setSize(18);
        $sheet->getStyle("B4")->getAlignment()->setWrapText(true);
        $this->applyHorizontalAlignment($sheet, "B4", Alignment::HORIZONTAL_CENTER);
        $this->applyVerticalAlignment($sheet, "B4", Alignment::VERTICAL_CENTER);

        for ($row = 4; $row <= 9; $row++) {
            $sheet->getRowDimension($row)->setRowHeight(18);
        }

        // Nomor & Pekerjaan
        $sheet->mergeCells("B10:H10");
        $sheet->setCellValue("B10", "Nomor               :  BA-PL/Jakarta/008/IX/2025");

        $sheet->mergeCells("I10:N10");
        $sheet->setCellValue("I10", "Pekerjaan :  Jasa Sewa GPS dan Dashcam");

        $sheet->getStyle("B10:N10")->getFont()->setSize(16);


        $sheet->getStyle("B10:H10")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_MEDIUM);
        $sheet->getStyle("I10:N10")->getBorders()->getOutline()->setBorderStyle(Border::BORDER_MEDIUM);
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
        $minRows  = count($data) > 2 ? count($data) : 2; // Minimum 2 baris body


        $ttdHeaderRow   = $startRow + 3;
        $ttdRow1        = $ttdHeaderRow + 1;
        $ttdRow2        = $ttdRow1 + 1;
        $ttdRow3        = $ttdRow2 + 1;

        $ttdRow4        = $ttdRow3 + 2;
        $ttdRow5        = $ttdRow4 + 1;
        $ttdRow6        = $ttdRow5 + 1;

        $column1 = $startCol;
        $column2 = chr(ord($column1) + 1);
        $column3 = chr(ord($column2) + 1);
        $column4 = chr(ord($column3) + 1);

        $sheet->setCellValue("{$column1}{$startRow}", "Pada hari ini, Rabu tanggal 01 bulan Oktober tahun Dua Ribu Dua Puluh Lima (2025)");

        $sheet->setCellValue("{$column1}{$ttdHeaderRow}", "Yang bertanda tangan dibawah ini : ");
        
        $sheet->setCellValue("{$column1}{$ttdRow1}", "I.");
        $sheet->setCellValue("{$column2}{$ttdRow1}", "Nama ");
        $sheet->setCellValue("{$column3}{$ttdRow1}", ":");
        $sheet->setCellValue("{$column4}{$ttdRow1}", "Muhammad Reza Kusuma ");

        $sheet->setCellValue("{$column2}{$ttdRow2}", "Jabatan ");
        $sheet->setCellValue("{$column3}{$ttdRow2}", ":");
        $sheet->setCellValue("{$column4}{$ttdRow2}", "Chief Operating Officer ");

        $sheet->setCellValue("{$column1}{$ttdRow3}", "Bertindak untuk dan atas nama PT Indi Daya Sistem selanjutnya mohon disebut sebagai Pihak Pertama");


        $sheet->setCellValue("{$column1}{$ttdRow4}", "II.");
        $sheet->setCellValue("{$column2}{$ttdRow4}", "Nama ");
        $sheet->setCellValue("{$column3}{$ttdRow4}", ":");
        $sheet->setCellValue("{$column4}{$ttdRow4}", "Bayu Riyadi ");

        $sheet->setCellValue("{$column2}{$ttdRow5}", "Jabatan ");
        $sheet->setCellValue("{$column3}{$ttdRow5}", ":");
        $sheet->setCellValue("{$column4}{$ttdRow5}", "Area Manager Jawa Bagian Barat ");

        $sheet->setCellValue("{$column1}{$ttdRow6}", "Bertindak untuk atas nama PT Patra Logistik selanjutnya mohon disebut sebagai Pihak Kedua");


        // BERDASARKAN
        $rowTableTitle      = 26;
        $rowTableSubTitle   = $rowTableTitle + 2;
        $rowTableHeader     = $rowTableSubTitle + 2;

        $sheet->mergeCells("{$startCol}{$rowTableTitle}:{$endCol}{$rowTableTitle}");
        $sheet->setCellValue("{$startCol}{$rowTableTitle}", "B E R D A S A R K A N");
        $sheet->getStyle("{$startCol}{$rowTableTitle}")->getFont()->setBold(true)->setUnderline(true)->setSize(14);
        $this->applyHorizontalAlignment($sheet, "{$startCol}{$rowTableTitle}", Alignment::HORIZONTAL_CENTER);

        $sheet->mergeCells("{$startCol}{$rowTableSubTitle}:{$endCol}{$rowTableSubTitle}");
        $sheet->setCellValue("{$startCol}{$rowTableSubTitle}", "Nomor Kontrak Perjanjian : KTR-857/PL000010/2024-S0");

        // HEADER TABLE
        $sheet->mergeCells("D{$rowTableHeader}:H{$rowTableHeader}");
        $sheet->mergeCells("I{$rowTableHeader}:J{$rowTableHeader}");
        $sheet->mergeCells("K{$rowTableHeader}:N{$rowTableHeader}");

        $sheet->setCellValue("B{$rowTableHeader}", "NO");
        $sheet->setCellValue("C{$rowTableHeader}", "URAIAN PEKERJAAN");
        $sheet->setCellValue("D{$rowTableHeader}", "TOTAL UNIT");
        $sheet->setCellValue("I{$rowTableHeader}", "HARGA SATUAN (Rp/Unit/Bulan)");
        $sheet->setCellValue("K{$rowTableHeader}", "TAGIHAN");

        $sheet->getRowDimension($rowTableHeader)->setRowHeight(27);
        $this->applyHeaderStyle($sheet, "{$startCol}{$rowTableHeader}:{$endCol}{$rowTableHeader}");

        // BODY TABLE
        $data = [
            [
                'name'          => 'GPS',
                'total_unit'    => 248,
                'harga_satuan'  => 350000,
                'tagihan'       => 86800000
            ],
            [
                'name'          => 'DASHCAM',
                'total_unit'    => 248,
                'harga_satuan'  => 566000,
                'tagihan'       => 140368000
            ],
        ];

        $bodyStart  = $rowTableHeader + 1;
        $bodyEnd    = $bodyStart + $minRows - 1;
        $footerRow  = $bodyEnd + 1;

        for ($i=0; $i<$minRows; $i++) {
            $currentRow = $bodyStart + $i;

            $sheet->mergeCells("D{$currentRow}:H{$currentRow}");
            $sheet->mergeCells("I{$currentRow}:J{$currentRow}");
            $sheet->mergeCells("K{$currentRow}:N{$currentRow}");

            if (isset($data[$i])) {
                $item = $data[$i];
                $sheet->setCellValue("B{$currentRow}", $i + 1);
                $sheet->setCellValue("C{$currentRow}", $item['name']);
                $sheet->setCellValue("D{$currentRow}", $item['total_unit']);
                $sheet->setCellValue("I{$currentRow}", $item['harga_satuan']);
                $sheet->setCellValue("K{$currentRow}", $item['tagihan']);
            }

            $this->applyHorizontalAlignment($sheet, "B{$currentRow}", Alignment::HORIZONTAL_CENTER);
            $this->applyHorizontalAlignment($sheet, "D{$currentRow}", Alignment::HORIZONTAL_RIGHT);
            $this->applyHorizontalAlignment($sheet, "I{$currentRow}", Alignment::HORIZONTAL_CENTER);
            $this->applyHorizontalAlignment($sheet, "K{$currentRow}", Alignment::HORIZONTAL_RIGHT);

            $sheet->getStyle("I{$currentRow}:K{$currentRow}")->getNumberFormat()->setFormatCode('"Rp"#,##0');

        }

        $this->applyBorder($sheet, "{$startCol}{$bodyStart}:{$endCol}{$bodyEnd}");

        // FOOTER TABLE
        $sheet->mergeCells("B{$footerRow}:J{$footerRow}");
        $sheet->mergeCells("K{$footerRow}:N{$footerRow}");

        $sheet->setCellValue("B{$footerRow}", "TOTAL");
        $sheet->setCellValue("K{$footerRow}", 227168000);

        $sheet->getStyle("K{$footerRow}")->getNumberFormat()->setFormatCode('"Rp"#,##0');
        $sheet->getStyle("B{$footerRow}:N{$footerRow}")->getFont()->setBold(true);

        $this->applyHorizontalAlignment($sheet, "B{$footerRow}", Alignment::HORIZONTAL_CENTER);
        $this->applyHorizontalAlignment($sheet, "K{$footerRow}", Alignment::HORIZONTAL_RIGHT);

        $this->applyBorder($sheet, "B{$footerRow}:N{$footerRow}");


        // FOOTER PAGE
        $footerPageRow  = $footerRow + 3;
        $ttdFooterRow1   = $footerPageRow + 6;
        $ttdFooterRow2   = $ttdFooterRow1 + 1;
        $ttdFooterRow3   = $ttdFooterRow2 + 9;
        $ttdFooterRow4   = $ttdFooterRow3 + 1;

        $sheet->mergeCells("C{$ttdFooterRow1}:F{$ttdFooterRow1}");
        $sheet->mergeCells("C{$ttdFooterRow2}:F{$ttdFooterRow2}");
        $sheet->mergeCells("C{$ttdFooterRow3}:F{$ttdFooterRow3}");
        $sheet->mergeCells("C{$ttdFooterRow4}:F{$ttdFooterRow4}");

        $sheet->mergeCells("J{$ttdFooterRow1}:M{$ttdFooterRow1}");
        $sheet->mergeCells("J{$ttdFooterRow2}:M{$ttdFooterRow2}");
        $sheet->mergeCells("J{$ttdFooterRow3}:M{$ttdFooterRow3}");
        $sheet->mergeCells("J{$ttdFooterRow4}:M{$ttdFooterRow4}");

        $sheet->setCellValue("C{$ttdFooterRow1}", "PIHAK PERTAMA");
        $sheet->setCellValue("C{$ttdFooterRow2}", " PT INDI DAYA SISTEM");
        $sheet->setCellValue("C{$ttdFooterRow3}", " Muhammad Reza Kusuma");
        $sheet->setCellValue("C{$ttdFooterRow4}", " Chief Operating Officer");

        $sheet->setCellValue("J{$ttdFooterRow1}", "PIHAK KEDUA");
        $sheet->setCellValue("J{$ttdFooterRow2}", "  PT PATRA LOGISTIK");
        $sheet->setCellValue("J{$ttdFooterRow3}", " Bayu Riyadi");
        $sheet->setCellValue("J{$ttdFooterRow4}", " Area Manager Jawa Bagian Barat");

        $sheet->getStyle("C{$ttdFooterRow1}:C{$ttdFooterRow2}")->getFont()->setBold(true);
        $sheet->getStyle("C{$ttdFooterRow3}")->getFont()->setBold(true)->setUnderline(true);

        $sheet->getStyle("J{$ttdFooterRow1}:J{$ttdFooterRow2}")->getFont()->setBold(true);
        $sheet->getStyle("J{$ttdFooterRow3}")->getFont()->setBold(true)->setUnderline(true);

        $this->applyHorizontalAlignment($sheet, "C{$ttdFooterRow1}:M{$ttdFooterRow4}", Alignment::HORIZONTAL_CENTER);


        $sheet->mergeCells("B{$footerPageRow}:N{$footerPageRow}");
        $sheet->setCellValue("B{$footerPageRow}", "Demikian Berita Acara ini dibuat untuk dipergunakan sebagaimana mestinya");


        $endRow = $ttdFooterRow4 + 1;
        $sheet->getStyle("B4:{$endCol}{$endRow}")
        ->getBorders()
        ->getOutline()
        ->setBorderStyle(Border::BORDER_MEDIUM);

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
                'startColor' => ['rgb' => 'efefef'],
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
            ->getFont()->setName('Arial');
    }

    /* =====================================================
     | COLUMN WIDTHS
     ===================================================== */
    public function columnWidths(): array
    {
        $data = [
            'A' => 2.18, 'B' => 5.50, 'C' => 29.17, 'D' => 1.00, 'E' => 11.67,
            'F' => 6.83, 'G' => 6.83, 'H' => 3.83, 'I' => 17.83, 'J' => 17.83,
            'K' => 9.00, 'L' => 9.33, 'M' => 10.33, 'N' => 11.50
        ];

        return $data;

        // return array_map(fn($v) => $v + 0.83, $data);
    }
}