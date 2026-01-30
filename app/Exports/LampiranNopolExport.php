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
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class LampiranNopolExport implements FromCollection, WithEvents, WithColumnWidths, WithTitle
{
    protected array $data;
    protected string $sheetName;

    protected int $lastRowKendaraanOperasi;
    protected int $lastRowPelepasanPerangkat;
    protected int $lastRowPemasanganBaru;

    public function __construct(array $data, string $sheetName = 'Lampiran_Nopol')
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
                    'endCol'    => 'H',
                    'startRow'  => 11,
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
        $sheet->mergeCells('A1:H10');

        // buat objek image
        $leftLogo = new Drawing();
        $leftLogo->setName('Image');
        $leftLogo->setDescription('Image A1:B4');
        $leftLogo->setPath(public_path('pertamina-patra-logistik.png'));

        // ukuran image (pixel)
        $leftLogo->setWidth(196.8);

        // anchor ke cell hasil merge
        $leftLogo->setCoordinates('A1');

        // optional: center-kan di merge cell
        $leftLogo->setOffsetX(15);
        $leftLogo->setOffsetY(20);

        $leftLogo->setWorksheet($sheet);

        // buat objek image
        $rightLogo = new Drawing();
        $rightLogo->setName('Image');
        $rightLogo->setDescription('Image G1:H4');
        $rightLogo->setPath(public_path('ids.png'));

        // ukuran image (pixel)
        $rightLogo->setWidth(195);

        // anchor ke cell hasil merge
        $rightLogo->setCoordinates('G1');

        // optional: center-kan di merge cell
        $rightLogo->setOffsetX(50);
        $rightLogo->setOffsetY(15);

        $rightLogo->setWorksheet($sheet);

        $richtextHeader = new RichText();
        $richtextHeader1 = $richtextHeader->createTextRun("LAMPIRAN");
        $richtextHeader1->getFont()->setBold(true)->setSize(15)->setName('Arial');
        $richtextHeader->createText("\n");

        $richtextHeader2 = $richtextHeader->createTextRun("PERIODE TANGGAL 1 - 30 SEPTEMBER 2025");
        $richtextHeader2->getFont()->setBold(true)->setSize(15)->setName('Arial');
        $richtextHeader->createText("\n");

        $richtextHeader3 = $richtextHeader->createTextRun("INTEGRATED TERMINAL PLUMPANG");
        $richtextHeader3->getFont()->setBold(true)->setSize(15)->setName('Arial');
        $richtextHeader->createText("\n\n");

        $richtextHeader4 = $richtextHeader->createTextRun("KENDARAAN OPERASI");
        $richtextHeader4->getFont()->setBold(true)->setSize(15)->setName('Arial');

        $sheet->setCellValue("A1", $richtextHeader);

        $sheet->getStyle("A1")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER);
     }
     private function renderHeaderPelepasanPerangkat($sheet, array $config)
     {
        $startCol = $config['startCol'];
        $endCol   = $config['endCol'];
        $startRow = $config['startRow'];
        $endRow   = $config['endRow'];

        $sheet->mergeCells("{$startCol}{$startRow}:{$endCol}{$endRow}");

        $richtextHeader = new RichText();
        $richtextHeader1 = $richtextHeader->createTextRun("LAMPIRAN");
        $richtextHeader1->getFont()->setBold(true)->setSize(15)->setName('Arial');
        $richtextHeader->createText("\n");

        $richtextHeader2 = $richtextHeader->createTextRun("PERIODE TANGGAL 1 - 30 SEPTEMBER 2025");
        $richtextHeader2->getFont()->setBold(true)->setSize(15)->setName('Arial');
        $richtextHeader->createText("\n");

        $richtextHeader3 = $richtextHeader->createTextRun("INTEGRATED TERMINAL PLUMPANG");
        $richtextHeader3->getFont()->setBold(true)->setSize(15)->setName('Arial');
        $richtextHeader->createText("\n\n");

        $richtextHeader4 = $richtextHeader->createTextRun("PELEPASAN PERANGKAT");
        $richtextHeader4->getFont()->setBold(true)->setSize(15)->setName('Arial');

        $sheet->setCellValue("{$startCol}{$startRow}", $richtextHeader);

        $sheet->getStyle("{$startCol}{$startRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER);
     }
     private function renderHeaderPemasanganBaru($sheet, array $config)
     {
        $startCol = $config['startCol'];
        $endCol   = $config['endCol'];
        $startRow = $config['startRow'];
        $endRow   = $config['endRow'];

        $sheet->mergeCells("{$startCol}{$startRow}:{$endCol}{$endRow}");

        $richtextHeader = new RichText();
        $richtextHeader1 = $richtextHeader->createTextRun("LAMPIRAN");
        $richtextHeader1->getFont()->setBold(true)->setSize(15)->setName('Arial');
        $richtextHeader->createText("\n");

        $richtextHeader2 = $richtextHeader->createTextRun("PERIODE TANGGAL 1 - 30 SEPTEMBER 2025");
        $richtextHeader2->getFont()->setBold(true)->setSize(15)->setName('Arial');
        $richtextHeader->createText("\n");

        $richtextHeader3 = $richtextHeader->createTextRun("INTEGRATED TERMINAL PLUMPANG");
        $richtextHeader3->getFont()->setBold(true)->setSize(15)->setName('Arial');
        $richtextHeader->createText("\n\n");

        $richtextHeader4 = $richtextHeader->createTextRun("PEMASANGAN BARU");
        $richtextHeader4->getFont()->setBold(true)->setSize(15)->setName('Arial');

        $sheet->setCellValue("{$startCol}{$startRow}", $richtextHeader);

        $sheet->getStyle("{$startCol}{$startRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)->setVertical(Alignment::VERTICAL_CENTER);
     }

     /* =====================================================
     | FOOTER
     ===================================================== */
     private function renderFooter($sheet)
     {
        // Tanda Tangan
        $signatureRow1 = $this->lastRowPemasanganBaru + 5;;
        $signatureRow2 = $signatureRow1 + 1;
        $signatureRow3 = $signatureRow2 + 10;
        $signatureRow4 = $signatureRow3 + 1;

        $sheet->mergeCells("B{$signatureRow1}:D{$signatureRow1}");
        $sheet->mergeCells("B{$signatureRow2}:D{$signatureRow2}");
        $sheet->mergeCells("B{$signatureRow3}:D{$signatureRow3}");
        $sheet->mergeCells("B{$signatureRow4}:D{$signatureRow4}");

        $sheet->mergeCells("F{$signatureRow1}:H{$signatureRow1}");
        $sheet->mergeCells("F{$signatureRow2}:H{$signatureRow2}");
        $sheet->mergeCells("F{$signatureRow3}:H{$signatureRow3}");
        $sheet->mergeCells("F{$signatureRow4}:H{$signatureRow4}");

        $sheet->setCellValue("B{$signatureRow1}", 'PIHAK PERTAMA');
        $sheet->setCellValue("B{$signatureRow2}", 'PT INDI DAYA SISTEM');
        $sheet->setCellValue("B{$signatureRow3}", 'Muhammad Reza Kusuma ');
        $sheet->setCellValue("B{$signatureRow4}", 'Chief Operating Officer');

        $sheet->setCellValue("F{$signatureRow1}", 'PIHAK KEDUA');
        $sheet->setCellValue("F{$signatureRow2}", 'PT PATRA LOGISTIK');
        $sheet->setCellValue("F{$signatureRow3}", 'Bayu Riyadi');
        $sheet->setCellValue("F{$signatureRow4}", 'Area Manager Jawa Bagian Barat');

        $sheet->getStyle("B{$signatureRow1}:B{$signatureRow4}")->getFont()->setBold(true);
        $sheet->getStyle("F{$signatureRow1}:F{$signatureRow4}")->getFont()->setBold(true);

        $sheet->getStyle("B{$signatureRow3}")->getFont()->setUnderline(true);
        $sheet->getStyle("F{$signatureRow3}")->getFont()->setUnderline(true);

        $sheet->getStyle("B{$signatureRow1}:H{$signatureRow4}")->getFont()->setSize(14)->setName("Times New Roman");
        $sheet->getStyle("B{$signatureRow1}:H{$signatureRow4}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
     }

     /* =====================================================
     | RENDER TABLE
     ===================================================== */
    private function renderTable($sheet, array $config)
    {

        $startColumn = $config['startCol'] ?? 'A';
        $endColumn   = $config['endCol'] ?? 'H';

        $startTableRow = $config['startRow'] ?? 11;

        $headerRow1 = $startTableRow;
        $headerRow2 = $startTableRow + 1;

        $bodyStartRow = $headerRow2 + 1;

        $startColumnIndex = Coordinate::columnIndexFromString($startColumn);

        $column1 = Coordinate::stringFromColumnIndex($startColumnIndex);        // No
        $column2 = Coordinate::stringFromColumnIndex($startColumnIndex + 1);    // No. Polisi
        $column3 = Coordinate::stringFromColumnIndex($startColumnIndex + 2);    // Merk Tipe
        $column4 = Coordinate::stringFromColumnIndex($startColumnIndex + 3);    // Unit Id GPS
        $column5 = Coordinate::stringFromColumnIndex($startColumnIndex + 4);    // IMEI GPS
        $column6 = Coordinate::stringFromColumnIndex($startColumnIndex + 5);    // Regional
        $column7 = Coordinate::stringFromColumnIndex($startColumnIndex + 6);    // FT/IT
        $column8 = Coordinate::stringFromColumnIndex($startColumnIndex + 7);    // Keterangan

        $columns = [$column1, $column2, $column3, $column4, $column5, $column6, $column7, $column8];
        $headers = [$headerRow1, $headerRow2];

        // KENDARAAN OPERASI
        $params = [
            'headers'       => $headers,
            'columns'       => $columns,
            'bodyStartRow'  => $bodyStartRow
        ];

        $this->sectionKendaraanOperasi($sheet, $params);

        // PELEPASAN PERANGKAT
        $headerPerangkatRow1 = $this->lastRowKendaraanOperasi + 12;
        $params = array_merge($params, [
            'headers'           => [$headerPerangkatRow1],
            'startRow'          => $this->lastRowKendaraanOperasi + 2,
            'bodyStartRow'      => $headerPerangkatRow1 + 1
        ]);

        $this->sectionPelepasanPerangkat($sheet, $params);

        // PEMASANGAN BARU
        $headerPemasanganRow1 = $this->lastRowPelepasanPerangkat + 12;
        $params = array_merge($params, [
            'headers'           => [$headerPemasanganRow1],
            'startRow'          => $this->lastRowPelepasanPerangkat + 2,
            'bodyStartRow'      => $headerPemasanganRow1 + 1
        ]);

        $this->sectionPemasanganBaru($sheet, $params);
    }

    private function sectionKendaraanOperasi($sheet, array $config)
    {
        $headerRow1 = $config['headers'][0];
        $headerRow2 = $config['headers'][1];

        $column1 = $config['columns'][0];
        $column2 = $config['columns'][1];
        $column3 = $config['columns'][2];
        $column4 = $config['columns'][3];
        $column5 = $config['columns'][4];
        $column6 = $config['columns'][5];
        $column7 = $config['columns'][6];
        $column8 = $config['columns'][7];

        $bodyStartRow = $config['bodyStartRow'];

        // HEADER
        $sheet->mergeCells("{$column1}{$headerRow1}:{$column1}{$headerRow2}")->setCellValue("{$column1}{$headerRow1}", "No");
        $sheet->mergeCells("{$column2}{$headerRow1}:{$column2}{$headerRow2}")->setCellValue("{$column2}{$headerRow1}", "No. Polisi");
        $sheet->mergeCells("{$column3}{$headerRow1}:{$column3}{$headerRow2}")->setCellValue("{$column3}{$headerRow1}", "Merk Tipe");
        $sheet->mergeCells("{$column4}{$headerRow1}:{$column4}{$headerRow2}")->setCellValue("{$column4}{$headerRow1}", "Unit Id GPS");
        $sheet->mergeCells("{$column5}{$headerRow1}:{$column5}{$headerRow2}")->setCellValue("{$column5}{$headerRow1}", "IMEI GPS");
        $sheet->mergeCells("{$column6}{$headerRow1}:{$column6}{$headerRow2}")->setCellValue("{$column6}{$headerRow1}", "Regional");
        $sheet->mergeCells("{$column7}{$headerRow1}:{$column7}{$headerRow2}")->setCellValue("{$column7}{$headerRow1}", "FT/IT");
        $sheet->mergeCells("{$column8}{$headerRow1}:{$column8}{$headerRow2}")->setCellValue("{$column8}{$headerRow1}", "Keterangan");


        $this->applyHeaderStyle($sheet, "{$column1}{$headerRow1}:{$column8}{$headerRow2}");

        // BODY
        $row = $bodyStartRow;
        for ($i=0; $i < 10; $i++) { 
            $sheet->setCellValue("{$column1}{$row}", $i + 1); // No
            $sheet->setCellValue("{$column2}{$row}", "B 9708 SEI"); // No. Polisi
            $sheet->setCellValue("{$column3}{$row}", "HINO"); // Merk Tipe
            $sheet->setCellValue("{$column4}{$row}", "R20010036"); // Unit Id GPS
            $sheet->setCellValueExplicit(
                "{$column5}{$row}",
                '860501044640270',
                DataType::TYPE_STRING
            );  // IMEI GPS
            $sheet->setCellValue("{$column6}{$row}", "RJBB"); // Regional
            $sheet->setCellValue("{$column7}{$row}", "IT Plumpang"); // FT/IT
            $sheet->setCellValue("{$column8}{$row}", "Operasi"); // Keterangan

            $row++;
        }

        $bodyEndRow = $row - 1;

        $notesRow = $bodyEndRow + 2;
        $this->lastRowKendaraanOperasi = $notesRow;

        $sheet->setCellValue("{$column2}{$notesRow}", "Note: Jumlah MT Operasi di bulan September sebanyak 248");
        $sheet->getStyle("{$column2}{$notesRow}")->getFont()->setBold(true)->setItalic(true)->setSize(10)->setName("Arial");

        $this->applyBodyStyle($sheet, "{$column1}{$bodyStartRow}:{$column8}{$bodyEndRow}");
        $sheet->getStyle("{$column5}{$bodyStartRow}:{$column5}{$bodyEndRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $sheet->getStyle("{$column1}{$bodyStartRow}:{$column1}{$bodyEndRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle("{$column6}{$bodyStartRow}:{$column8}{$bodyEndRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
    }
    private function sectionPelepasanPerangkat($sheet, array $config)
    {
        $headerRow1 = $config['headers'][0];

        $column1 = $config['columns'][0];
        $column2 = $config['columns'][1];
        $column3 = $config['columns'][2];
        $column4 = $config['columns'][3];
        $column5 = $config['columns'][4];
        $column6 = $config['columns'][5];
        $column7 = $config['columns'][6];
        $column8 = $config['columns'][7];

        $bodyStartRow = $config['bodyStartRow'];

        $configHeader = [
            'startCol'  => $column1,
            'endCol'    => $column8,
            'startRow'  => $config['startRow'],
            'endRow'    => $config['startRow'] + 9
        ];

        $this->renderHeaderPelepasanPerangkat($sheet, $configHeader);
        $sheet->getStyle("{$column1}{$headerRow1}:{$column8}{$headerRow1}")->getFont()->setName("Arial");

        // HEADER
        $sheet->mergeCells("{$column3}{$headerRow1}:{$column5}{$headerRow1}");
        $sheet->mergeCells("{$column7}{$headerRow1}:{$column8}{$headerRow1}");

        $sheet->setCellValue("{$column1}{$headerRow1}", "No.");                 // No.
        $sheet->setCellValue("{$column2}{$headerRow1}", "Nopol");               // Nopol
        $sheet->setCellValue("{$column3}{$headerRow1}", "Keterangan");          // Keterangan
        $sheet->setCellValue("{$column6}{$headerRow1}", "FT/IT");               // FT/IT
        $sheet->setCellValue("{$column7}{$headerRow1}", "Tanggal Pencopotan "); // Tanggal Pencopotan

        $this->applyHeaderStyle($sheet, "{$column1}{$headerRow1}:{$column8}{$headerRow1}");

        // BODY
        $row = $bodyStartRow;
        for ($i=0; $i < 1; $i++) { 

            $sheet->mergeCells("{$column3}{$row}:{$column5}{$row}");
            $sheet->mergeCells("{$column7}{$row}:{$column8}{$row}");

            $sheet->setCellValue("{$column1}{$row}", null);    // No
            $sheet->setCellValue("{$column2}{$row}", null);    // Nopol
            $sheet->setCellValue("{$column3}{$row}", null);    // Keterangan
            $sheet->setCellValue("{$column6}{$row}", null);    // FT/IT
            $sheet->setCellValue("{$column7}{$row}", null);    // Tanggal Pencopotan

            $row++;
        }

        $bodyEndRow = $row - 1;

        $notesRow = $bodyEndRow + 2;
        $this->lastRowPelepasanPerangkat = $notesRow + 4;

        $sheet->setCellValue("{$column2}{$notesRow}", "Note: Tidak ada MT pelepasan baru di bulan September");
        $sheet->getStyle("{$column2}{$notesRow}")->getFont()->setBold(true)->setItalic(true)->setSize(10)->setName("Arial");

        $this->applyBodyStyle($sheet, "{$column1}{$bodyStartRow}:{$column8}{$bodyEndRow}");
        $sheet->getStyle("{$column1}{$bodyStartRow}:{$column8}{$bodyEndRow}")->getFont()->setName("Arial");

        // $sheet->getStyle("{$column5}{$bodyStartRow}:{$column5}{$bodyEndRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        // $sheet->getStyle("{$column1}{$bodyStartRow}:{$column1}{$bodyEndRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        // $sheet->getStyle("{$column6}{$bodyStartRow}:{$column8}{$bodyEndRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

    }
    private function sectionPemasanganBaru($sheet, array $config)
    {
        $headerRow1 = $config['headers'][0];

        $column1 = $config['columns'][0];
        $column2 = $config['columns'][1];
        $column3 = $config['columns'][2];
        $column4 = $config['columns'][3];
        $column5 = $config['columns'][4];
        $column6 = $config['columns'][5];
        $column7 = $config['columns'][6];
        $column8 = $config['columns'][7];

        $bodyStartRow = $config['bodyStartRow'];

        $configHeader = [
            'startCol'  => $column1,
            'endCol'    => $column8,
            'startRow'  => $config['startRow'],
            'endRow'    => $config['startRow'] + 9
        ];

        $this->renderHeaderPemasanganBaru($sheet, $configHeader);

        // HEADER
        $sheet->mergeCells("{$column3}{$headerRow1}:{$column5}{$headerRow1}");
        $sheet->mergeCells("{$column7}{$headerRow1}:{$column8}{$headerRow1}");

        $sheet->setCellValue("{$column1}{$headerRow1}", "No.");                 // No.
        $sheet->setCellValue("{$column2}{$headerRow1}", "Nopol");               // Nopol
        $sheet->setCellValue("{$column3}{$headerRow1}", "Keterangan");          // Keterangan
        $sheet->setCellValue("{$column6}{$headerRow1}", "FT/IT");               // FT/IT
        $sheet->setCellValue("{$column7}{$headerRow1}", "Tanggal Pemasangan "); // Tanggal Pemasangan

        $this->applyHeaderStyle($sheet, "{$column1}{$headerRow1}:{$column8}{$headerRow1}");
        $sheet->getStyle("{$column1}{$headerRow1}:{$column8}{$headerRow1}")->getFont()->setName("Arial");

        // BODY
        $row = $bodyStartRow;
        for ($i=0; $i < 1; $i++) { 

            $sheet->mergeCells("{$column3}{$row}:{$column5}{$row}");
            $sheet->mergeCells("{$column7}{$row}:{$column8}{$row}");

            $sheet->setCellValue("{$column1}{$row}", null);    // No
            $sheet->setCellValue("{$column2}{$row}", null);    // Nopol
            $sheet->setCellValue("{$column3}{$row}", null);    // Keterangan
            $sheet->setCellValue("{$column6}{$row}", null);    // FT/IT
            $sheet->setCellValue("{$column7}{$row}", null);    // Tanggal Pemasangan

            $row++;
        }

        $bodyEndRow = $row - 1;

        $notesRow = $bodyEndRow + 2;
        $this->lastRowPemasanganBaru = $notesRow;

        $sheet->setCellValue("{$column2}{$notesRow}", "Note: Tidak ada MT pemasangan device baru di bulan September");
        $sheet->getStyle("{$column2}{$notesRow}")->getFont()->setBold(true)->setItalic(true)->setSize(10)->setName("Arial");

        $this->applyBodyStyle($sheet, "{$column1}{$bodyStartRow}:{$column8}{$bodyEndRow}");
        $sheet->getStyle("{$column1}{$bodyStartRow}:{$column8}{$bodyEndRow}")->getFont()->setName("Arial");

        // $sheet->getStyle("{$column5}{$bodyStartRow}:{$column5}{$bodyEndRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        // $sheet->getStyle("{$column1}{$bodyStartRow}:{$column1}{$bodyEndRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        // $sheet->getStyle("{$column6}{$bodyStartRow}:{$column8}{$bodyEndRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

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
            ->getFont()->setName("Times New Roman");
    }

    /* =====================================================
     | COLUMN WIDTHS
     ===================================================== */
    public function columnWidths(): array
    {
        $data = [
            'A' => 5.33, 'B' => 18.83, 'C' => 14.33, 'D' => 15.33, 'E' => 19.67,
            'F' => 13.67, 'G' => 16, 'H' => 16.83, 'I' => 13.17, 'J' => 7.83
        ];

        return array_map(fn($v) => $v + 0.83, $data);
    }
}