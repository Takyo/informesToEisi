<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;


/**
 * TODO:
 *      el orden de meter el array sea como en la cabecera
 *      meter que se pueda poner en negrita facilmente
 *      show muestre una tabla html incluso con los colores indicados y que sea ¿bootstrap?
 *      opcion para anclar filas o columnas
 *      log
 *      warning que se quede guardado el warning en ele log cuando el numero de cabecera no coincide con el el cuerpo
 *      comentar bien con dockblocker
 *
 * clase que genera un excel sencillo de informes
 * donde la primera fila es la cabecera de esta
 */
class Este
{
    public $objExcel;
    public $log;
    protected $filename;
    public $headers;
    public $headerStyle;
    protected $datas;
    protected $hojas;
    // protected $removeCharToSpace;
    protected $assoc; // si se tienen en cuenta las keys del array como headers
    protected $rango; // rango de letras (ancho), se autocalcula

    const SIMPLE = 'sim';
    const MULTI_ARRAY = 'multi';
    const MULTI_ARRAY_ASOC = 'multi_asoc';

    /**
     * TODO: opciones que sean opcionales y un default
     */
    function __construct(Spreadsheet $objExcel = null, array $opt = null)
    {
        $this->objExcel = $objExcel;
        $this->filename = $opt['filename']; // nombre del archivo
        $this->headerStyle = [
            'font' => [
                'bold' => true,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'color' => array('argb' => 'FFA0A0A0'),
            ],
        ];
        $this->orderFromHeader = $opt['order_from_header']; // el orden lo marca el encabezado (lento) o segun datos pasados (fast)

        $this->log = array('init' => time(), 'Inicio');
    }

    /**
     * metadatos que saldrán en la primera hoja
     * @param $inicio default true, indica si se va a la ultima hoja o al principio
     */
    public function info(array $info, $inicio = true)
    {
        //informe titulo del informe
        //creado datetime
    }

    /**
     * dice si el array pasado es asociativo
     * NOTA: multiarray no lo tiene en cuenta
     */
    private function is_assoc(array $array)
    {
        return count(array_filter(array_keys($array), 'is_string')) > 0;
    }

    /**
     * dice si el array el simple o multiArray
     */
    private function is_multiArray(array $array)
    {
        return count($array) != count($array, COUNT_RECURSIVE);
    }

    /**
     * dice si el array es un multiarray  asociativo
     */
    private function is_multiArray_assoc(array $array)
    {
        if ($this->is_multiArray($array)){
            foreach($array as $key => $elem) {
                // print_r($elem); echo $key;
                // echo gettype($key);
                // if(!$this->is_assoc($elem) || gettype($key) == 'integer') {
                if(gettype($key) == 'integer') {
                    return false;
                }
            }
        } else {
            return false;
        }

        return true;
    }

    /**
     * genera la cabecera o cabeceras del informe
     * TODO:
     *       - comprobar que sea un row o multiples con el nombre de la hoja como indice
     *       - ver si realmente hace falta esta funcion
     */
    public function setHeader(array $header, array $style = null)
    {
        $this->headers = $header;
        $styleArray = [
            'font' => [
                'bold' => true,
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
            ],
            'borders' => [
                'top' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startColor' => [
                    'argb' => 'FFA0A0A0',
                ],
                'endColor' => [
                    'argb' => 'FFFFFFFF',
                ],
            ],
        ];
       // $spreadsheet->getActiveSheet()->getStyle('B3:B7')->applyFromArray($styleArray);
    }


    /**
     * Aplica estilo al header
     *
     * @param array   $headerStyle    array en formato stilo (ver documentacion)
     * @param int     $hoja           número de hoja a aplicar (defecto 0)
     * @return void
     */
    public function setHeaderStyle(array $headerStyle = null, $hoja = 0)
    {

        if (!isset($headerStyle) || !is_array($headerStyle)) {
            $headerStyle = [
                'font' => [
                    'bold' => true
                ],
                'fill' => [
                    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                    'color' => array('argb' => 'FFA0A0A0'),
                ],
            ];
        }

        $this->headerStyle = $headerStyle;

        $active = $this->objExcel->setActiveSheetIndex($hoja);

        $max = $active->getHighestColumn();

        $active->getStyle('A1:'.$max.'1')
               ->applyFromArray($headerStyle);

        $active->freezePane('A2');

        $this->addLog('Aplicando estilo a la cabecera '.$hoja);
    }

    /**
    * Añade una nueva row a la tabla
    */
    public function addRow(array $row, $hoja = null)
    {
    }

    /**
     * datos de todas las hojas
     * NOTA:
     *      el array tiene que ir con un formato espefico
     *
     * @param array $datos
     * @param boolean $arrayAsociativo  (opcional). Indica si el nombre de las keys del array $datos es la cabecera de las columnas
     * @return void
     */
    public function datos(array $datos, bool $arrayAsociativo=null)
    {
    }

    /**
     * Añade una hoja al futuro excel
     *
     * NOTA:
     *      - Siempre que se indique una cabecera está se pondrá
     *      - El header buscará en el array asociativo
     *      - En caso de no indicarle la cabecera se mirara si el array es asociativo y la cabecera
     *          será las keys del primer row de este
     *      - Si $datos no fuera un array asociativo no tendrá header está hoja
     *
     * @param $titulo     hoja
     * @param $datos
     * @param $cabecera   (opcional), si no se indica cabecera o false rellena datos
     *                              si true toma como cabecera las keys del array $datos
     */
    public function addHoja($titulo, array $dato, array $header = null)
    {
        $hoja = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($this->objExcel, $titulo);

        // CASO 1: se le pasa el header como parámetro
        if (!empty($header)) {
            $hoja->fromArray($header);
        }

        // CASO 2: no se le pasa el header y el multiarray $datos es asociativo
        // cogemos solo las keys del primer registro
        else if ($this->is_multiArray($dato) && $this->is_assoc($dato[0])) {
            $hoja->fromArray(array_keys($dato[0]));
        }

        // CASO 3: no se le pasa el header y el array $datos son asociativos
        else if ($this->is_assoc($dato)) {
            $header = $hoja->fromArray(array_keys($dato));
        }

        // $typeHeader = $this->array_is($header);

        // inserta por asociacion según el orden del header
        if (!empty($header)) {
            $letra = 'A';
            foreach ($header as $key => $val) {
                $col = array_chunk(array_column($dato, $val), 1);
                $hoja->fromArray($col, null, $letra.'2');
                $letra++;
            }

            // NOTA: No se pueden aplicar estilos a la hoja si antes no se le añade al excel
            $this->objExcel->addSheet($hoja);
            $this->setHeaderStyle(null, $this->objExcel->getSheetCount() - 1);
        }
        // sin header, sin estilos
        else {
            $hoja->fromArray($dato,null,'A1');
            // NOTA: No se pueden aplicar estilos a la hoja si antes no se le añade al excel
            $this->objExcel->addSheet($hoja);
        }
    }

    /**
     * añade log
     * @param string $log
     */
    private function addLog($log)
    {
        array_push($this->log, $log);
    }

    /**
     * imprime por pantalla el log
     */
    public function printLog()
    {
        echo '<pre>';
        foreach ($this->log as $log) {
            echo $log.PHP_EOL;
        }
        echo '</pre>';
    }

    // TODO: descarga el log.txt
    public function DownloadLog()
    {
    }

    /**
     * muestra el array montado
     * TODO: sin terminar, mostrar una tabla cutre
     */
    public function show()
    {
        echo "<pre>";
        echo "CABECERA";
        print_r($this->headers);
        echo "\nDATOS";
        print_r($this->datas);
    }

    /**
     * dice de que tipo array es: simple, multi array simple, multi array asociativo
     *
     * @params array $array
     *
     * @return string 'simple' | 'multi' | 'multi_asoc'
     */
    private function array_is(array $array)
    {
        if ($this->is_multiArray_assoc($array)) {
            return $this::MULTI_ARRAY_ASOC;
        } else if ($this->is_multiArray($array)) {  // multi simple
            return $this::MULTI_ARRAY;
        } else {
            return $this::SIMPLE;
        }
    }

    /**
     *  TODO:
     *   comprobar headers = null
     *   tengo que mirar cuando es un array simple
     *   mirar que cada elemento de los datos se relacione con la cabecera
     *   mirar ordenamiento?
     *   opcion que los datos se pongan como vengan, para ir mas rapido
     *   dato que no este en cabecera, dato que no se pone
     */
    public function generate($datos, $headers)
    {
        $this->remove();
        $typeDatos = $this->array_is($datos);
        $typeHeader = $this->array_is($headers);

        echo "datos: $typeDatos <br>";
        echo "header: $typeHeader <br>";

        // genera X paginas
        if ($typeDatos == $this::MULTI_ARRAY_ASOC) {

            if ($typeHeader == $this::SIMPLE) {
                foreach ($datos as $titulo => $hoja) {
                    $this->addHoja($titulo, $hoja, $headers);
                }
            } else if ($typeHeader == $this::MULTI_ARRAY) {
                $h = 0;
                foreach ($datos as $titulo => $hoja) {
                    $head = (array_key_exists($h, $headers)) ? $headers[$h] : NULL;
                    $this->addHoja($titulo, $hoja, $head);
                    $h++;
                }
            }

            // busca el titulo en las cabeceras
            // TODO: generar un warning si no encuentra cabecera
            else if ($typeHeader == $this::MULTI_ARRAY_ASOC) {
                foreach ($datos as $titulo => $hoja) {
                    $head = (array_key_exists($titulo, $headers)) ? $headers[$titulo] : NULL;
                    $this->addHoja($titulo, $hoja, $head);
                }
            }
        }

        // si es un simple array o multi array, será una sola página
        else {

            if ($typeHeader == $this::SIMPLE) {
                $this->addHoja('Informe', $datos, $headers);
            }

            // solo tiene como cabecera el primer header
            else if ($typeHeader == $this::MULTI_ARRAY) {
                $this->addHoja('Informe', $datos, $headers[0]);
            }

            // tiene en cuenta como cabecera el primer elemento
            else if ($typeHeader == $this::MULTI_ARRAY_ASOC) {
                foreach ($headers as $titulo => $head) {
                    $this->addHoja($titulo, $datos, $head);
                    break;
                }
            }
        }

        $this->addLog("Generado Excel");
        $this->addLog('Fin');
        array_push($this->log, ['end' => time()]);
        // $this->log = array('end' => time());
    }

    /**
     * manda el excel al navegador
     *
     * @param $format (opcional) default: xls.   Indica el tipo de documento xls/xlsx
     */
    public function download($format=null)
    {

        $this->objExcel->setActiveSheetIndex(0);

        switch($format) {
            case 'xlsx':
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                $format = 'xlsx';
            break;
            case 'xls':
            case null:
            default:
                header('Content-Type: application/vnd.ms-excel');
                $format = 'xls';
            break;
        }

        header("Content-Disposition: attachment;filename='$this->filename.$format'");
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $writer = IOFactory::createWriter($this->objExcel, 'Xls');
        $writer->save('php://output');

        $this->addLog("Mandando Excel " . $this->filename . $format);

    }

    /**
     * destruye la hoja indicada por su posicion o por el name
     * TODO: mirar si da error cuando eliminas algo que no existe o el cero dos veces
     */
    public function remove($search = 0)
    {

        $index = 0;
        switch(gettype($search)) {
            case 'integer':
                $index = $search;
            break;
            case 'string':
                $index = $this->objExcel->getIndex(
                    $this->objExcel->getSheetByName($search)
                );
            break;
            default:
                return;
        }

        $this->objExcel->removeSheetByIndex($index);

        $this->addLog("Eliminada hoja $index");
    }
}


// $excel = new ExcelToEisi();
// [
//     'titulo' => 'informe',
//     'meta' => $info,
//     'datos' =>
//     [
//         'titulo' => 'hoja1',
//         'cabecera' => $cabecera,
//         'datos' => $datos,
//     ],[
//         'titulo' => 'hoja2',
//         'cabecera' => $cabecera,
//         'datos' => $datos,
//     ],[
//         'titulo' => 'hoja3',
//         'cabecera' => $cabecera,
//         'datos' => $datos,
//     ]
// ]
?>
