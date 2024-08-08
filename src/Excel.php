<?php

namespace Jeanp\Jexcel;

use Maatwebsite\Excel\Facades\Excel as FacadesExcel;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

/*
  - PHP EXCEL LARAVEL 9
  - composer require psr/simple-cache:^1.0 maatwebsite/excel
  - php artisan vendor:publish --provider="Maatwebsite\Excel\ExcelServiceProvider" --tag=config
  - providers(config/app.php) ->
  -     Maatwebsite\Excel\ExcelServiceProvider::class,
        App\Providers\PHPExcelMacroServiceProvider::class,
  - aliases -> 'Excel' => Maatwebsite\Excel\Facades\Excel::class,
  - Agregar archivo Providers/PHPExcelMacroServiceProvider.php
 */

class Excel
{

    private $event = null;
    private $sheet = null;
    private $cell = null;
    private $column = null;
    private $row = null;

    public function __construct($event)
    {
        $this->event = $event;
        $this->sheet = $this->event->sheet;
    }

    public static function import($instaceModel, $path = "")
    {
        FacadesExcel::import($instaceModel, $path);
    }

    public static function store($model, $data, $filename, $disk = 'public'): string
    {
        $path = "exports/" . $filename;

        FacadesExcel::store(new $model($data), $path, $disk);

        return $path;
    }

    public static function download($model, $data, $filename)
    {

        return FacadesExcel::download(new $model($data), $filename);
    }


    /**
     * En el excel se muestra en centímetros y aquí son pulgadas, convertir de centrímetros a pulgadas
     */
    public function setPageMargin($top, $right, $bottom, $left)
    {

        $factor = 2.54;

        $this->sheet->getPageMargins()->setTop($top / $factor);
        $this->sheet->getPageMargins()->setRight($right / $factor);
        $this->sheet->getPageMargins()->setBottom($bottom / $factor);
        $this->sheet->getPageMargins()->setLeft($left / $factor);
    }

    public function landscape()
    {
        $this->sheet->getPageSetup()
            ->setPaperSize(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LEGAL);
    }

    public function toArray($file)
    {

        return FacadesExcel::toArray([], $file); //Obtener data del excel en Array
    }


    /**
     * @param string $text Título principal que llevará el Excel Celda A1
     */
    public function title(string $text)
    {
        $this->cell('A1', $text)->fontSize(16)->bold();

        return $this;
    }

    /**
     * @param string $cell Celda en la que se insertará el texto - @example A1, A2, B7
     * @param string $text Texto que se escribirá en la celda
     */
    public function subtitle(string $cell, string $text)
    {
        $this->cell($cell, $text)->fontSize(14);

        return $this;
    }

    /**
     * @param array $arr Array de strings columnas
     * @param number $row Número de fila
     */
    public function thead($arr = [], $row = 5)
    {

        $letterInit = 1; //A
        $this->row = $row;

        foreach ($arr as $text) {
            $letter = Coordinate::stringFromColumnIndex($letterInit);
            $this->cell($letter . $row, $text)->align('center')->background('D9D9D9')->bold();

            $letterInit++;
        }

        $firstLetter = Coordinate::stringFromColumnIndex(1);
        $lastLetter = Coordinate::stringFromColumnIndex($letterInit - 1);

        $this->borders($firstLetter . $row . ":" . $lastLetter . $row);

        return $this;
    }

    /**
     * @param int $number Número de columna a convertir en texto. Exam. 1 -> A, 2 -> B
     */
    public function columnLetter(int $number)
    {
        return Coordinate::stringFromColumnIndex($number);
    }

    /**
     * @param string $cell Celda en la cuál se escribirá
     * @param mixed  $value Texto que se mostrará en la celda
     * @param mixed  $type  Formato de texto [decimal|date]
     */
    public function cell(string $cell, $value = '', $type = null)
    {
        $this->cell = $cell;

        if ($value) {

            if ($type == 'string') {
                $this->sheet->setCellValueExplicit($this->cell, $value,  DataType::TYPE_STRING);
            } else if ($type == 'decimal') {
                $this->sheet->setCellValue($this->cell, $value);
                $this->formatNumber($this->cell);
            } else {
                $this->sheet->setCellValue($this->cell, $value);
            }
        }

        return $this;
    }



    /**
     * @param string $cell Celda en la cuál se combinará
     */
    public function merge(string $cell)
    {

        $this->cell = $this->cell . ':' . $cell;

        $this->sheet->mergeCells($this->cell);

        return $this;
    }

    /**
     * @param string $letter Letra de la columna
     * @param number $width  Ancho de la columna
     */
    public function column(string $letter, $width = null)
    {
        $this->column = $letter;

        if ($width) {
            $this->sheet->getColumnDimension(strtoupper($letter))->setWidth($width);
        }

        return $this;
    }

    /**
     * @param array $arr Lista de columnas con su width
     */
    public function columns($arr = [])
    {

        foreach ($arr as $letter => $width) {
            $this->column($letter, $width);
        }
    }
    /**
     * @param int $row Número de fila
     * @param int $height Tamaño vertical de fila
     */
    public function height($row, $height = null)
    {
        $this->row = $row;

        if ($height) {
            $this->sheet->getRowDimension($row)->setRowHeight($height);
        }

        return $this;
    }

    /**
     * Height de acuerdo al contenido
     */
    public function autoHeight()
    {
        $this->sheet->getRowDimension($this->row)->setRowHeight(-1);

        return $this;
    }

    /**
     * @param string $color Código Hexadecimal
     */
    public function background(string $color)
    {
        $this->style($this->cell, [
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'color' => ['argb' => str_replace("#", "", $color)]
            ]
        ]);

        return $this;
    }

    /**
     * @param string $color Código Hexadecimal
     */
    public function color(string $color)
    {
        $this->style($this->cell, [
            'font' => [
                'color' => ['argb' => str_replace("#", "", $color)]
            ]
        ]);

        return $this;
    }

    /**
     * @param number $num Tamaño de fuente
     */
    public function fontSize($num = 11)
    {

        $this->style($this->cell, [
            'font' => [
                'size' => $num
            ],
        ]);

        return $this;
    }

    public function bold()
    {
        $this->style($this->cell, [
            'font' => [
                'bold' => true
            ],
        ]);

        return $this;
    }

    /**
     * @param left|center|right $positionx Posición en la que se situará el texto horizontalmente
     * @param top|bottom $positiony Posición en la que se situará el texto verticalmente
     */


    public function align(string $positionx, string|null $positiony = null)
    {


        if ($positionx == 'center') {
            $this->sheet->horizontalAlign($this->cell, \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        } else if ($positionx == 'right') {
            $this->sheet->horizontalAlign($this->cell, \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
        }

        if ($positiony) {

            if ($positiony == 'top') {
                $this->sheet->verticalAlign($this->cell, \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);
            } else if ($positiony == 'bottom') {
                $this->sheet->verticalAlign($this->cell, \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM);
            } else if ($positiony == 'middle' || $positiony == 'center') {
                $this->sheet->verticalAlign($this->cell, \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
            } else if ($positiony == 'justify') {
                $this->sheet->verticalAlign($this->cell, \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_JUSTIFY);
            }
        }

        return $this;
    }

    public function alignJustify()
    {
        $this->sheet->horizontalAlign($this->cell, \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $this->sheet->verticalAlign($this->cell, \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $this->wrap();

        return $this;
    }

    public function style($range, $styles = [])
    {
        $this->sheet->styleCells($range, $styles);

        return $this;
    }

    /**
     * @param string $range Rango de celdas a la cuál se aplicarán los borders
     * @param string $type allBorders|top|bottom|right|left
     */
    public function borders($range, $type = "allBorders")
    {
        $this->style($range, [
            'borders' => [
                $type => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['argb' => '000000'],
                ],
            ]
        ]);

        return $this;
    }

    public function formatNumber($range)
    {
        $this->style(
            $range,
            [
                'numberFormat' => [
                    'formatCode' =>  \PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1
                ]
            ]
        );
    }


    public function wrap(bool $isWrap = true)
    {
        $this->style($this->cell, [
            'wrap' => $isWrap,
        ]);

        if ($isWrap) {
            $this->sheet->wrapText($this->cell);
        }

        return $this;
    }
}
