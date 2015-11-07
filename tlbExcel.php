<?php

class tlbExcel {

    const DATA_TYPE_VIEW_MODEL = 'viewModel';
    const DATA_TYPE_CAPTION = 'caption';
    const DATA_TYPE_TABLE = 'table';

    //the PHPExcel object
    public $libPath = 'ext.phpexcel.Classes.PHPExcel'; //the path to the PHP excel lib
    public static $objPHPExcel = null;
    public static $activeSheet = null;
    //Document properties
    public $creator = 'Uldis Nelsons';
    public $title = null;
    public $subject = 'Subject';
    public $description = '';
    public $category = '';
    public $lastModifiedBy = '';
    public $keywords = '';
    public $sheetTitle = '';
    public $legal = 'PHPExcel generator http://phpexcel.codeplex.com/ - EExcelView Yii extension http://yiiframework.com/extension/eexcelview/ - Adaptation by A. Bennouna http://tellibus.com';
    public $landscapeDisplay = false;
    public $A4 = false;
    public $RTL = false;
    public $pageFooterText = '&RPage &P of &N';
    //config
    public $offset = 0;
    public $autoWidth = true;
    public $exportType = 'Excel2007';
    public $disablePaging = true;
    public $filename = null; //export FileName
    public $stream = true; //stream to browser
    //options
    public $decimalSeparator = '.';
    public $thousandsSeparator = ',';
    public $displayZeros = false;
    public $zeroPlaceholder = '-';
    public $border_style;
    public $borderColor = '000000';
    public $bgColor = 'FFFFFF';
    public $textColor = '000000';
    public $rowHeight = 15;
    public $headerBorderColor = '000000';
    public $headerBgColor = 'CCCCCC';
    public $headerTextColor = '000000';
    public $headerHeight = 20;
    public $footerBorderColor = '000000';
    public $footerBgColor = 'FFFFCC';
    public $footerTextColor = '0000FF';
    public $footerHeight = 20;
    public $zoomScale = 100;
    public static $fill_solid;
    public static $papersize_A4;
    public static $orientation_landscape;
    public static $horizontal_center;
    public static $horizontal_right;
    public static $vertical_center;
    public static $style = [];
    public static $headerStyle = [];
    public static $footerStyle = [];
    public static $captionStyle = [];
    //buttons config
    public $exportButtonsCSS = 'summary';
    public $exportButtons = ['Excel2007'];
    public $exportText = 'Export to: ';
    //callbacks
    public $onRenderHeaderCell = null;
    public $onRenderDataCell = null;
    public $onRenderFooterCell = null;
    public $company;
    //mime types used for streaming
    public $mimeTypes = [
        'Excel5' => [
            'Content-type' => 'application/vnd.ms-excel',
            'extension' => 'xls',
            'caption' => 'Excel(*.xls)',
        ],
        'Excel2007' => [
            'Content-type' => 'application/vnd.ms-excel',
            'extension' => 'xlsx',
            'caption' => 'Excel(*.xlsx)',
        ],
        'PDF' => [
            'Content-type' => 'application/pdf',
            'extension' => 'pdf',
            'caption' => 'PDF(*.pdf)',
        ],
        'HTML' => [
            'Content-type' => 'text/html',
            'extension' => 'html',
            'caption' => 'HTML(*.html)',
        ],
        'CSV' => [
            'Content-type' => 'application/csv',
            'extension' => 'csv',
            'caption' => 'CSV(*.csv)',
        ]
    ];
    public $data;

    public function init() {

        $lib = Yii::getPathOfAlias($this->libPath) . '.php';
        if (!file_exists($lib)) {
            $this->grid_mode = 'grid';
            Yii::log("PHP Excel lib not found($lib). Export disabled !", CLogger::LEVEL_WARNING, 'EExcelview');
        }

        if (!isset($this->title)) {
            $this->title = Yii::app()->getController()->getPageTitle();
        }


        //Autoload fix
        spl_autoload_unregister(['YiiBase', 'autoload']);
        Yii::import($this->libPath, true);

        // Get here some PHPExcel constants in order to use them elsewhere
        self::$papersize_A4 = PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4;
        self::$orientation_landscape = PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE;
        self::$fill_solid = PHPExcel_Style_Fill::FILL_SOLID;
        if (!isset($this->border_style)) {
            $this->border_style = PHPExcel_Style_Border::BORDER_THIN;
        }
        self::$horizontal_center = PHPExcel_Style_Alignment::HORIZONTAL_CENTER;
        self::$horizontal_right = PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
        self::$vertical_center = PHPExcel_Style_Alignment::VERTICAL_CENTER;

        spl_autoload_register(['YiiBase', 'autoload']);

        // Creating a workbook
        self::$objPHPExcel = new PHPExcel();
        self::$activeSheet = self::$objPHPExcel->getActiveSheet();

        // Set some basic document properties
        if ($this->landscapeDisplay) {
            self::$activeSheet->getPageSetup()->setOrientation(self::$orientation_landscape);
        }

        if ($this->A4) {
            self::$activeSheet->getPageSetup()->setPaperSize(self::$papersize_A4);
        }

        if ($this->RTL) {
            self::$activeSheet->setRightToLeft(true);
        }

        self::$objPHPExcel->getProperties()
                ->setTitle($this->title)
                ->setCreator($this->creator)
                ->setSubject($this->subject)
                ->setDescription($this->description . ' // ' . $this->legal)
                ->setCategory($this->category)
                ->setLastModifiedBy($this->lastModifiedBy)
                ->setKeywords($this->keywords);

        // Initialize styles that will be used later
        self::$style = [
            'borders' => [
                'allborders' => [
                    'style' => $this->border_style,
                    'color' => ['rgb' => $this->borderColor],
                ],
            ],
            'fill' => [
                'type' => self::$fill_solid,
                'color' => ['rgb' => $this->bgColor],
            ],
            'font' => [
                //'bold' => false,
                'color' => ['rgb' => $this->textColor],
            ]
        ];
        self::$headerStyle = [
            'borders' => [
                'allborders' => [
                    'style' => $this->border_style,
                    'color' => ['rgb' => $this->headerBorderColor],
                ],
            ],
            'fill' => [
                'type' => self::$fill_solid,
                'color' => ['rgb' => $this->headerBgColor],
            ],
            'font' => [
                'bold' => true,
                'color' => ['rgb' => $this->headerTextColor],
            ]
        ];

        self::$captionStyle = [
            'borders' => [
                'allborders' => [
                    'style' => '',
                //'color' => ['rgb' => '000000'],
                ],
            ],
            'fill' => [
                'type' => self::$fill_solid,
                'color' => ['rgb' => 'FFFFFF'],
            ],
            'font' => [
                'bold' => true,
                'size' => 16,
                'color' => ['rgb' => '000000'],
            ]
        ];
    }

    public function renderHeader() {



//        if (!empty($this->caption)) {
//
//            self::$activeSheet->mergeCells($this->columnName(1) . (string) $this->offset . ':' . $this->columnName($this->caption_columns) . (string) $this->offset);
//            $cell = self::$activeSheet->setCellValue($this->columnName(1) . (string) $this->offset, $this->caption, true);
//            self::$activeSheet->getRowDimension($this->offset)->setRowHeight(20);
//            $caption = self::$activeSheet->getStyle($this->columnName(1) . (string) $this->offset . ':' . $this->columnName(5) . (string) $this->offset);
//            $caption->applyFromArray(self::$captionStyle);
//            $this->offset += 1;
//        }

        $a = 0;

        // Format the header row
//        $header = self::$activeSheet->getStyle($this->columnName(1) . (string) ($this->offset + 1) . ':' . $this->columnName($a) . (string) ($this->offset + 1));
//        $header->getAlignment()
//                ->setHorizontal(self::$horizontal_center)
//                ->setVertical(self::$vertical_center);
//        $header->applyFromArray(self::$headerStyle);
//        self::$activeSheet->getRowDimension(1)->setRowHeight($this->headerHeight);
//
//        $this->offset += 1;
    }

    public function renderView($data) {
        $x = $data['x'];
        $y = $data['y'];
        $model = $data['model'];

        self::$activeSheet->getColumnDimension($this->columnName($x))->setAutoSize(true);
        self::$activeSheet->getColumnDimension($this->columnName($x + 1))->setAutoSize(true);

        foreach ($data['attributes'] as $attribute) {
            if (isset($attribute['value'])) {
                $value = $attribute['value'];
            } else {
                $value = $model->$attribute['name'];
            }
            $label = $model->getAttributeLabel($attribute['name']);

            $cell = self::$activeSheet->setCellValue($this->columnName($x) . $y, $label, true);
            $cell = self::$activeSheet->setCellValue($this->columnName($x + 1) . $y, $value, true);
            $y ++;
        }

        return [
            'x' => $x,
            'y' => $y,
        ];
    }

    public function renderTable($data) {
        $xBase = $x = $data['x'];
        $y = $data['y'];
        $header = $data['header'];
        $rows = $data['rows'];

        /**
         * header
         */
        foreach ($header as $name => $label) {
            $cord = $this->columnName($x) . $y;
            $cell = self::$activeSheet->setCellValue($cord, $label, true);

            $cellStyle = self::$activeSheet->getStyle($cord);
            $cellStyle->applyFromArray(self::$headerStyle);

            self::$activeSheet->getColumnDimension($this->columnName($x))->setAutoSize(true);

            $x ++;
        }
        $y ++;

        /**
         * body
         */
        foreach ($rows as $row) {
            $x = $xBase;
            foreach ($row as $fieldName => $fieldValue) {
                if (is_null($row)) {
                    $x ++;
                    continue;
                }
                if (!is_array($fieldValue)) {
                    $cell = self::$activeSheet->setCellValue($this->columnName($x) . $y, $fieldValue, true);
                    $x ++;
                    continue;
                }
                if (isset($fieldValue['type'])) {
                    $fieldValue['x'] = $x;
                    $fieldValue['y'] = $y;
                    $r = $this->renderDataBlock($fieldValue);
                    $x = $r['x'];
                    $y = $r['y'];
                    $x ++;
                }
            }
            $y ++;
        }

        return [
            'x' => $x,
            'y' => $y,
        ];
    }

    public function renderCaption($data) {
        $x = $data['x'];
        $y = $data['y'];
        $width = $data['width'];
        $caption = $data['caption'];

        $cell1Cord = $this->columnName($x) . (string) $y;
        $cell2Cord = $this->columnName((int) $x + (int) $width) . (string) $y;
        $mCord = $cell1Cord . ':' . $cell2Cord;
        self::$activeSheet->mergeCells($mCord);
        $cell = self::$activeSheet->setCellValue($cell1Cord, $caption, true);
        //self::$activeSheet->getRowDimension($x)->setRowHeight(20);
        $cellStyle = self::$activeSheet->getStyle($mCord);
        $cellStyle->applyFromArray(self::$captionStyle);

        $y ++;
        return [
            'x' => $x,
            'y' => $y,
        ];
    }

    public function renderDataBlock($data) {


        switch ($data['type']) {
            case self::DATA_TYPE_VIEW_MODEL:
                return $this->renderView($data);
                break;
            case self::DATA_TYPE_CAPTION:
                return $this->renderCaption($data);
                break;
            case self::DATA_TYPE_TABLE:
                return $this->renderTable($data);
                break;

            default:
                throw new Exception('Undefinet type:' . $data['type']);
                break;
        }
    }

    public function run() {
        $this->renderHeader();
        foreach ($this->data as $data) {
            $this->renderDataBlock($data);
        }
        // Set some additional properties
        self::$activeSheet
                ->setTitle($this->sheetTitle)
                ->getSheetView()->setZoomScale($this->zoomScale);

        self::$activeSheet->getHeaderFooter()
                ->setOddHeader('&C' . $this->sheetTitle)
                ->setOddFooter('&L&B' . self::$objPHPExcel->getProperties()->getTitle() . $this->pageFooterText);
        /**
         * printing settings
         */
//            self::$activeSheet->getPageSetup()
//                    ->setPrintArea('A1:' . $this->columnName(count($this->columns)) . ($this->offset + $row + 2))
//                    ->setFitToWidth();
        //create writer for saving
        $objWriter = PHPExcel_IOFactory::createWriter(self::$objPHPExcel, $this->exportType);
        if (!$this->stream) {
            $objWriter->save($this->filename);
        } else {
            //output to browser
            if (!$this->filename) {
                $this->filename = $this->title;
            }
            $this->cleanOutput();
            header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
            header('Pragma: public');
            header('Content-type: ' . $this->mimeTypes[$this->exportType]['Content-type']);
            header('Content-Disposition: attachment; filename="' . $this->filename . '.' . $this->mimeTypes[$this->exportType]['extension'] . '"');
            header('Cache-Control: max-age=0');
            $objWriter->save('php://output');
            Yii::app()->end();
        }
    }

    /**
     * Returns the corresponding Excel column.(Abdul Rehman from yii forum)
     * 
     * @param int $index
     * @return string
     */
    public function columnName($index) {
        --$index;
        if (($index >= 0) && ($index < 26)) {
            return chr(ord('A') + $index);
        } else if ($index > 25) {
            return ($this->columnName($index / 26)) . ($this->columnName($index % 26 + 1));
        } else {
            throw new Exception("Invalid Column # " . ($index + 1));
        }
    }

    /**
     * Performs cleaning on mutliple levels.
     * 
     * From le_top @ yiiframework.com
     * 
     */
    private static function cleanOutput() {
        for ($level = ob_get_level(); $level > 0; --$level) {
            @ob_end_clean();
        }
    }

}
