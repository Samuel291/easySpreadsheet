<?php

    namespace Samuel291;

    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Style\Border;
    use PhpOffice\PhpSpreadsheet\Style\Fill;
    use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

    class EasySpreadsheet
    {
        /** @var Spreadsheet */
        private $spreadsheet;
        private $sheet;
        private $header;
        private $content;
        private $items;
        private $counter;

        public function __construct()
        {
            $this->spreadsheet = new Spreadsheet();
            $this->counter = 0;
        }

        public function startSheet(string $title = null): self
        {
            $this->sheet = ($this->counter == 0) ? $this->spreadsheet->getActiveSheet()->setTitle($title) : $this->spreadsheet->addSheet((new Worksheet($this->spreadsheet, $title)), $this->counter);
            $this->header = array();
            $this->items = array();
            $this->content = array();
            $this->counter++;
            return $this;
        }

        public function setHeader(string $title, ?array $settings = null): self
        {
            $data = new \stdClass();
            $data->title = $title;
            $data->settings = $settings;

            array_push($this->header, $data);

            return $this;
        }

        private function header(): self
        {
            $defaultStyle = [
                'font' => [
                    'size' => 7,
                    'bold' => true
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => ['argb' => '000000'],
                    ],
                ],
            ];

            foreach ($this->header as $i => $content) {
                $content = (object)$content;
                $this->sheet->setCellValue($this->getLetter($i) . '1', $content->title);
                $this->sheet->getStyle($this->getLetter($i) . '1')->applyFromArray($defaultStyle);
                if (!empty($content->settings)) {
                    $settings = (object)$content->settings;
                    if (!empty($settings->alignment)) {
                        $this->sheet->getStyle($this->getLetter($i) . '1')->getAlignment()->setWrapText(true)->setHorizontal('center');
                    }
                    if (!empty($settings->numberFormat)) {
                        $this->sheet->getStyle($this->getLetter($i))->getNumberFormat()->setFormatCode($settings->numberFormat);
                    }
                    if (!empty($settings->cellColor)) {
                        $this->cellColor($this->getLetter($i) . '1', $settings->cellColor);
                    }
                    if (!empty($settings->columnDimension)) {
                        $this->sheet->getColumnDimension($this->getLetter($i))->setWidth($settings->columnDimension);
                    } else {
                        $this->sheet->getColumnDimension($this->getLetter($i))->setAutoSize(true);
                    }
                    if (!empty($settings->merge)) {
                        $this->merge($i, 1, $settings->merge, $defaultStyle);
                    }
                } else {
                    $this->sheet->getColumnDimension($this->getLetter($i))->setAutoSize(true);
                }
            }
            return $this;
        }

        public function setItem(string $item, ?array $settings = null): self
        {
            $data = new \stdClass();
            $data->item = $item;
            $data->settings = $settings;

            array_push($this->items, $data);

            return $this;
        }

        public function setContent(?array $settings = null): self
        {
            array_push($this->content, $this->items);
            $this->items = [];
            $l = count($this->content) + 1;

            if (!empty($settings)) {
                $settings = (object)$settings;
                $cells = $this->getLetter(array_key_first($this->header)) . $l . ':' . $this->getLetter(array_key_last($this->header)) . $l;
                if (!empty($settings->cellColor)) {
                    $this->cellColor($cells, $settings->cellColor);
                }
                if (!empty($settings->fontColor)) {
                    $this->fontColor($cells, $settings->fontColor);
                }
            }
            return $this;
        }

        public function content(): self
        {
            $l = 1;
            $defaultStyle = [
                'font' => [
                    'size' => 7
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => ['argb' => '000000'],
                    ],
                ],
            ];
            foreach ($this->content as $items) {
                $l++;
                $c = 0;
                foreach ($items as $item) {
                    $item = (object)$item;
                    $this->sheet->setCellValue($this->getLetter($c) . $l, $item->item);
                    $this->sheet->getStyle($this->getLetter($c) . $l)->applyFromArray($defaultStyle);
                    if (!empty($item->settings)) {
                        $settings = (object)$item->settings;
                        if (!empty($settings->cellColor)) {
                            $this->cellColor($this->getLetter($c) . $l, $settings->cellColor);
                        }
                        if (!empty($settings->fontColor)) {
                            $this->fontColor($this->getLetter($c) . $l, $settings->fontColor);
                        }
                        if (!empty($settings->merge)) {
                            $this->merge($c, $l, $settings->merge, $defaultStyle);
                        }
                        if (!empty($settings->fontWeight)) {
                            $this->sheet->getStyle($this->getLetter($c) . $l)->getFont()->setBold(true);
                        }
                        if (!empty($settings->alignment)) {
                            $this->sheet->getStyle($this->getLetter($c) . $l)->getAlignment()->setWrapText(true)->setHorizontal('center');
                        }
                        if (!empty($settings->comment)) {
                            $this->sheet->getComment($this->getLetter($c) . $l)->getText()->createTextRun($settings->comment);
                        }
                    }
                    $c++;
                }
            }
            return $this;
        }

        public function render()
        {
            $this->header();
            $this->content();
            return $this;
        }

        public function save(string $filename, ?string $path = null): void
        {
            $writer = new Xlsx($this->spreadsheet);
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment; filename="' . urlencode($this->removeAccents($filename) . '.xlsx') . '"');
            $writer->save(!empty($path) ? $path . '/' . $this->removeAccents($filename) . '.xlsx' : 'php://output');
        }

        private function getLetter(int $n): string
        {
            $a = range('A', 'Z');
            return ($n <= 25) ? $a[$n] : $a[((int)($n / 26) - 1)] . $a[($n % 26)];
        }

        private function fontColor(string $cells, string $color): void
        {
            $this->sheet->getStyle($cells)->getFont()->getColor()->setARGB($color);
        }

        private function cellColor(string $cells, string $color): void
        {
            $this->sheet->getStyle($cells)->getFill()->setFillType(Fill::FILL_SOLID);
            $this->sheet->getStyle($cells)->getFill()->getStartColor()->setARGB($color);
            $this->sheet->getStyle($cells)->getFill()->getEndColor()->setARGB($color);
        }

        private function merge(int $c, int $l, string $direction, $defaultStyle): void
        {
            $nCells = filter_var($direction, FILTER_SANITIZE_NUMBER_INT);
            $d = (strpos($direction, 'left')) ? $c - $nCells : $c + $nCells;
            $cells = strpos($direction, 'left') ? $this->getLetter($d) . $l . ':' . $this->getLetter($c) . $l : $this->getLetter($c) . $l . ':' . $this->getLetter($d) . $l;
            $this->sheet->getStyle($cells)->applyFromArray($defaultStyle);
            $this->sheet->mergeCells($cells);
        }

        private function removeAccents(string $string)
        {
            $acentos = 'ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûýýþÿŔŕ';
            $sem_acentos = 'aaaaaaaceeeeiiiidnoooooouuuuybsaaaaaaaceeeeiiiidnoooooouuuyybyRr';
            $string = strtr(utf8_decode($string), utf8_decode($acentos), $sem_acentos);
            return utf8_decode($string);
        }
    }