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

class KegagalanPenormalanExport implements FromCollection, WithEvents, WithColumnWidths, WithTitle
{
    protected array $data;
    protected string $sheetName;
    protected string $lastColumn;

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


                $totalDate = count($this->data['item']);

                $startColIndex = Coordinate::columnIndexFromString('E');
                $endColIndex   = $startColIndex + $totalDate - 1;
                
                $this->lastColumn = Coordinate::stringFromColumnIndex($endColIndex + 2);

                // Left Logo
                $sheet->mergeCells("D1:D3");
                // buat objek image
                $drawing = new Drawing();
                $drawing->setName('Image');
                $drawing->setDescription('Image D1:D3');
                $drawing->setPath(public_path('ids.png'));

                // ukuran image (pixel)
                $drawing->setWidth(180);

                // anchor ke cell hasil merge
                $drawing->setCoordinates('D1');

                // optional: center-kan di merge cell
                $drawing->setOffsetX(0);
                $drawing->setOffsetY(0);

                $drawing->setWorksheet($sheet);

                // Right Logo
                $sheet->mergeCells("I1:I3");
                // buat objek image
                $rightLogo = new Drawing();
                $rightLogo->setName('Image');
                $rightLogo->setDescription('Image I1:I3');
                $rightLogo->setPath(public_path('pertamina-patra-logistik.png'));

                // ukuran image (pixel)
                $rightLogo->setWidth(180);

                // anchor ke cell hasil merge
                $rightLogo->setCoordinates("{$this->lastColumn}1");

                // optional: center-kan di merge cell
                $rightLogo->setOffsetX(60);
                $rightLogo->setOffsetY(0);

                $rightLogo->setWorksheet($sheet);

                $this->renderItemTable($sheet);
                $this->renderHeader($sheet);
                $this->renderFooter($sheet);
                $this->renderInformation($sheet);

                foreach ([4, 5, 6, 7] as $row) {
                    $sheet->getStyle("B{$row}")
                        ->getAlignment()
                        ->setHorizontal(Alignment::HORIZONTAL_LEFT)
                        ->setVertical(Alignment::VERTICAL_CENTER);
                }

                // $sheet->setShowGridlines(false);
                // $sheet->getRowDimension(6)->setRowHeight(26);
            }
        ];
    }

    /* =====================================================
     | HEADER
     ===================================================== */
    private function renderHeader($sheet): void
    {
        $sheet->setCellValue('D4', 'LAMPIRAN 1 & 2');
        $sheet->setCellValue('D5', 'REKAPITULASI KEGAGALAN & PENORMALAN OPERASI GPS - DASHCAM');
        $sheet->setCellValue('D6', 'PERIODE    : ' . $this->data['periode'] ?? 'SEPTEMBER 2025');
        $sheet->setCellValue('D7', 'LOKASI       : ' . $this->data['site'] ?? 'IT Jakarta');
        
        $sheet->getStyle('D4')->getFont()->setBold(true)->setSize(14);
        $sheet->getStyle('D5:D7')->getFont()->setBold(true)->setSize(12);
    }

    /* =====================================================
     | FOOTER
     ===================================================== */
     private function renderFooter($sheet): void
     {
        $sheet->setCellValue('D28', 'Disiapkan Oleh,');
        $sheet->setCellValue('D29', 'PT Indi Daya Sistem');
        $sheet->setCellValue('D30', 'Supervisor Services');
        $sheet->setCellValue('D37', 'Dimas Nafidin');

        $lastCol = $this->lastColumn;
        $sheet->setCellValue("{$lastCol}28", 'Disetujui Oleh,');
        $sheet->setCellValue("{$lastCol}29", 'PT  Patra Logistik');
        $sheet->setCellValue("{$lastCol}30", 'Area Manager Jawa Bagian Barat');
        $sheet->setCellValue("{$lastCol}37", 'Bayu Riyadi');
        
        $sheet->getStyle("D28:{$lastCol}37")->getFont()->setSize(14);
        $sheet->getStyle("D29:D37")->getFont()->setBold(true);
        $sheet->getStyle("{$lastCol}29:{$lastCol}37")->getFont()->setBold(true);
     }

     /* =====================================================
     | TABLE
     ===================================================== */
    private function renderItemTable($sheet): void
    {
        $data = $this->data;

        $totalDate = count($data['item']);

        $startColIndex = Coordinate::columnIndexFromString('E');
        $endColIndex   = $startColIndex + $totalDate - 1;

        $startColDate = Coordinate::stringFromColumnIndex($startColIndex);
        $endColDate   = Coordinate::stringFromColumnIndex($endColIndex);

        $columnTotal = Coordinate::stringFromColumnIndex($endColIndex + 1);
        $columnDesc  = Coordinate::stringFromColumnIndex($endColIndex + 2);

        // ===== HEADER =====
        $sheet->mergeCells('C9:C10')->setCellValue('C9', 'No');
        $sheet->mergeCells('D9:D10')->setCellValue('D9', 'Jenis Kegiatan');
        $sheet->mergeCells("{$startColDate}9:{$endColDate}9")->setCellValue("{$startColDate}9", 'Jan-26');

        // tanggal
        foreach ($data['item'] as $i => $item) {
            $col = Coordinate::stringFromColumnIndex($startColIndex + $i);

            $day = (int) date('j', strtotime($item['date'])); // j = no leading zero

            $sheet->setCellValue("{$col}10", $day);
            $sheet->getColumnDimension($col)->setWidth(5.5);
        }

        $sheet->mergeCells("{$columnTotal}9:{$columnTotal}10")->setCellValue("{$columnTotal}9", 'Total');
        $sheet->mergeCells("{$columnDesc}9:{$columnDesc}10")->setCellValue("{$columnDesc}9", 'Keterangan');

        $sheet->getStyle("C9:{$columnDesc}10")->getFont()->setBold(true)->setSize(11);
        $this->applyHeaderStyle($sheet, "C9:{$columnDesc}10");

        // ===== BODY =====
        $listJenisKegiatan = [
            ['value' => 'header', 'label' => 'Kegagalan Operasi'],
            ['value' => 'count_blackout', 'label' => 'Jumlah Black Out dengan Kejadian >30 Menit (per Unit Mobil)'],
            ['value' => 'value_kpi', 'label' => 'Nilai KPI'],
            ['value' => 'spacer', 'label' => null],
            ['value' => 'header', 'label' => 'Penormalan Operasi'],
            ['value' => 'normaly_time', 'label' => 'Waktu penormalan operasi pasca kegagalan operasi dengan waktu >1 jam'],
            ['value' => 'value_kpi_normal', 'label' => 'Nilai KPI'],
        ];

        $row = 11;
        $no  = 1;

        foreach ($listJenisKegiatan as $type) {

            if ($type['value'] === 'spacer') {
                $row++;
                continue;
            }

            $sheet->setCellValue("D{$row}", $type['label']);

            if ($type['value'] === 'header') {
                $sheet->setCellValue("C{$row}", $no++);
                $sheet->getStyle("C{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER); 
                $sheet->getStyle("D{$row}:{$columnDesc}{$row}")->applyFromArray([
                    'font' => ['bold' => true],
                    'fill' => [
                        'fillType' => Fill::FILL_SOLID,
                        'startColor' => ['rgb' => 'FFCC01'],
                    ],
                ]);
                $row++;
                continue;
            }

            // isi data per tanggal
            foreach ($data['item'] as $i => $item) {
                $col = Coordinate::stringFromColumnIndex($startColIndex + $i);
                $sheet->setCellValue("{$col}{$row}", $item[$type['value']] ?? 0);

                if (
                    $type['value'] === 'value_kpi' ||
                    $type['value'] === 'value_kpi_normal'
                ) {
                    $value = $item[$type['value']] ?? 0;

                    $sheet->setCellValueExplicit(
                        "{$col}{$row}",
                        $value / 100,
                        DataType::TYPE_NUMERIC
                    );

                    $sheet->getStyle("{$col}{$row}")
                        ->getNumberFormat()
                        ->setFormatCode('0%');
                }
            }

            $row++;
        }

        $row = $row - 1;

        $this->lastColumn = $columnDesc;
        
        $sheet->setCellValue("{$columnTotal}12", $this->data['total_blackout']);
        $sheet->setCellValue("{$columnTotal}13", $this->data['total_kpi_kegagalan']);

        $sheet->setCellValue("{$columnTotal}16", $this->data['total_time_penormalan']);
        $sheet->setCellValue("{$columnTotal}17", $this->data['total_kpi_penormalan']);

        $sheet->setCellValueExplicit("{$columnTotal}13", $this->data['total_kpi_kegagalan'] / 100, DataType::TYPE_NUMERIC);
        $sheet->getStyle("{$columnTotal}13")->getNumberFormat()->setFormatCode('0%');

        $sheet->setCellValueExplicit("{$columnTotal}17", $this->data['total_kpi_penormalan'] / 100, DataType::TYPE_NUMERIC);
        $sheet->getStyle("{$columnTotal}17")->getNumberFormat()->setFormatCode('0%');

        $sheet->setCellValue("{$columnDesc}12", "Angka perhari dari laporan harian");
        $sheet->setCellValue("{$columnDesc}16", "Angka dari kegagalan operasi");

        $this->applyBackgroundColor($sheet, "{$columnTotal}12:{$columnTotal}13", 'ffff00');
        $this->applyBackgroundColor($sheet, "{$columnTotal}16:{$columnTotal}17", 'ffff00');

        $this->applyBodyStyle($sheet, "C11:{$columnDesc}{$row}");

        $sheet->getStyle("C14:{$columnDesc}14")->getBorders()->getAllBorders()
            ->setBorderStyle(Border::BORDER_NONE);

        $sheet->getStyle("C14:{$columnDesc}14")->getBorders()->getOutline()
            ->setBorderStyle(Border::BORDER_THIN);

        $sheet->getStyle("E11:{$columnTotal}{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->getStyle("{$columnTotal}11:{$columnTotal}{$row}")->getFont()->setBold(true);

        $sheet->getStyle("C11:C{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->getColumnDimension($columnTotal)->setWidth(7.17);
        $sheet->getColumnDimension($columnDesc)->setWidth(33.5);
    }

    /* =====================================================
    | TABLE
    ===================================================== */
    private function renderInformation($sheet)
    {
        $sheet->getRowDimension(21)->setRowHeight(115);
        $sheet->setCellValue("D19", "Tata Cara Perhitungan KPI:");
        $sheet->getStyle("D19")->getFont()->setBold(true)->setUnderline(true);

        // 1. Kegagalan Operasi
        $sheet->setCellValue("D20", "1. Kegagalan Operasi");
        $richtextKegagalan = new RichText();
        $richtextKegagalan->createText("Setiap kelipatan 1 kasus mengurangi 5%");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("Performance Score Parameter KPI");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("0 kasus = 110 %");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("1 kasus = 95 %");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("2 Kasus = 90%");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("3 Kasus = 85%");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("dst.");

        $sheet->getStyle("D21")->getAlignment()->setWrapText(true);

        $sheet->setCellValue("D21", $richtextKegagalan);

        // Remaks
        $sheet->mergeCells("E21:H21");
        $sheet->getStyle("E21")->getAlignment()->setWrapText(true);

        $richtextRemaks = new RichText();
        $boldText = $richtextRemaks->createTextRun("Remaks :");
        $boldText->getFont()->setBold(true);
        $richtextRemaks->createText("\n");
        $richtextRemaks->createText("Kegagalan dan penormalan perangkat dicatat harian berdasarkan jumlah unit Mobil Tangki yang mengalami kegagalan, dengan satuan pengambilan per unit mobil.");

        $sheet->setCellValue("E21", $richtextRemaks);

        // 2. Penormalan Operasi
        $sheet->setCellValue("D23", "2. Penormalan Operasi");
        $richtextKegagalan = new RichText();
        // Kriteria :
        // Setiap kelipatan 1 kasus penormalam
        // yang melebihi 1 jam, akan mengurangi
        // 5% Performance Score Parameter KPI
        // 0 kasus = 110 %
        // 1 kasus = 95 %
        // 2 Kasus = 90%
        // 3 Kasus = 85%
        // dst.
        $richtextKegagalan->createText("Kriteria :");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("Setiap kelipatan 1 kasus penormalan");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("yang melebihi 1 jam, akan mengurangi");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("5% Perfomance Score Parameter KPI");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("0 kasus = 110 %");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("1 kasus = 95 %");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("2 Kasus = 90%");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("3 Kasus = 85%");
        $richtextKegagalan->createText("\n");
        $richtextKegagalan->createText("dst.");

        $sheet->getStyle("D24")->getAlignment()->setWrapText(true);

        $sheet->setCellValue("D24", $richtextKegagalan);

        $sheet->getStyle("D24:E24")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
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
        $data = [
            'A' => 3.83, 'B' => 0.64, 'C' => 2.67, 'D' => 56.17, 'E' => 4.00
        ];

        return array_map(fn($v) => $v + 0.83, $data);
    }
}