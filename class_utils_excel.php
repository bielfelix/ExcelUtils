<?php
include_once('framework/componentes/PhpSpreadsheet-1.8.0/src/Bootstrap.php'); // Versão compatível com PHP 5.6

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class ExcelUtils
{
    private static $spreadsheet;
    private static $sheet;
    private static $writer;

    private static $prefixName;
    private static $column = array();
    private static $field = array();
    private static $dados = array();

    private static $contador_auto = 1; //contador utlizado para a funcão nativa da classe na field  => contador_auto

   

    /**
     * Importa os dados de um arquivo Excel para um array
     *
     * @param string $file Caminho completo do arquivo Excel
     * @param int $sheetIndex Índice da planilha a ser importada (padrão é 0, a primeira planilha)
     * @return array|false Retorna um array com os dados da planilha ou false em caso de erro
     */

     public static function importar($file, $sheetIndex = 0)
     {
         try {
             // Carrega o arquivo Excel
             $spreadsheet = IOFactory::load($file);
             $sheet = $spreadsheet->getSheet($sheetIndex);
 
             $data = array();
             $highestRow = $sheet->getHighestRow();
             $highestColumn = $sheet->getHighestColumn();
 
             $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);
 
             for ($row = 1; $row <= $highestRow; ++$row) {
                 $rowData = array();
                 for ($col = 1; $col <= $highestColumnIndex; ++$col) {
                     $cellValue = $sheet->getCellByColumnAndRow($col, $row)->getValue();
                     $rowData[] = $cellValue;
                 }
                 $data[] = $rowData;
             }
 
             return $data;
         } catch (Exception $e) {
             // Trate qualquer exceção que possa ocorrer ao carregar o arquivo Excel
             return false;
         }
     }

    /**
     * Method para gerar planinhas excel em diferentes formatos.
     *
     * @param array $colunas  Array com nome das colunas, campos e funções de callback.
     *
     * @param array $dados  Os dados podem vim direto da query ou dentro de um array.
     *
     * @param string $arqName  Será o prefixo do nome do arquivo.
     *
     * @param string $tipo  Extensão do arquivo a ser gerada (xls, xlsx e csv).
     *
     * @param bool $dadosBD  true como default, indica se esta vindo direto de uma query e false quando for dentro de um array.
     *
     * @param string $delimitadorCSV   Defina o delimitador apropriado para o CSV, como padrão está o ponto e vírgula ( ; ).
     * 
     * @return void Retorna o arquivo para download.
     */

    public static function gerar($colunas, $dados, $arqName, $tipo, $dadosBD = true, $delimitadorCSV = ';')
    {
        self::$spreadsheet = new Spreadsheet();
        self::$sheet = self::$spreadsheet->getActiveSheet();

        self::$prefixName = $arqName.'_'.date('YmdHis');

        self::setArrayDados($dados, $dadosBD);
        
        self::setColumnAndField($colunas);

        $flush = ob_end_flush(); // limpa o buffer para poder montar a planilha

        self::setHeader();
        self::setBody();

        self::build($tipo, $delimitadorCSV);
    }

    private static function setHeader()
    {
        $colunaIndex = 1; // Iniciar a partir da coluna 1 (A)
        foreach (self::$column as $coluna) {
            $colunaLetra = self::getColunaLetra($colunaIndex);
            self::$sheet->setCellValue($colunaLetra . '1', $coluna);
            $colunaIndex++;
        }
    }

    private static function setBody()
    {
        $row = 2; // Iniciar a partir da segunda linha
        foreach (self::$dados as $item) {
            $colunaIndex = 1; // Iniciar a partir da coluna 1 (A)

            foreach (self::$field as $campo) {
                $colunaLetra = self::getColunaLetra($colunaIndex);
                $count_conteudo = intval(count($campo));
                if($count_conteudo > 1){
                    $array_conteudo = array();
                    for ($i = 0; $i <= ($count_conteudo - 1); $i++) {

                        if($campo[$i][1] != null){ $conteudo = self::setFormartDados($item, $campo[$i]);
                        }else{ $conteudo = $item[$campo[$i][0]]; }
                        
                        array_push($array_conteudo, mb_convert_encoding($conteudo, 'UTF-8', 'ISO-8859-1'));
                    }
                    self::$sheet->setCellValue($colunaLetra . $row, implode(' - ', $array_conteudo));
                }else{
                    
                    if($campo[0][1] != null){ $conteudo = self::setFormartDados($item, $campo[0]);
                    }else{ $conteudo = $item[$campo[0][0]]; }

                    self::$sheet->setCellValue($colunaLetra . $row, mb_convert_encoding($conteudo, 'UTF-8', 'ISO-8859-1'));
                }
                $colunaIndex++;
            }
            $row++;
            self::$contador_auto++;
        }
    }

    private static function build($tipo, $delimitadorCSV)
    {
        ob_clean();
        ob_start();

        switch ($tipo) {
            case 'xls':
                self::$writer = new Xls(self::$spreadsheet);
                header('Content-Type: application/vnd.ms-excel');
                header('Content-Disposition: attachment;filename="'.self::$prefixName.'.xls"');
                break;
            case 'xlsx':
                self::$writer = new Xlsx(self::$spreadsheet);
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Content-Disposition: attachment;filename="'.self::$prefixName.'.xlsx"');
                break;
            case 'csv':
                self::$writer = new Csv(self::$spreadsheet);
                self::$writer->setDelimiter($delimitadorCSV); // Defina o delimitador apropriado para o CSV, como ponto e vírgula (;) ou vírgula (,)
                self::$writer->setEnclosure(''); // Não use aspas para delimitar campos (opcional)
                self::$writer->setUseBOM(true); // Defina o BOM para garantir a codificação correta

                header('Content-Type: text/csv; charset=ISO-8859-1');
                header('Content-Disposition: attachment;filename="'.self::$prefixName.'.csv"');
                break;
            default:
                throw new Exception("Tipo de arquivo inválido.");
        }

        ob_end_clean();
        ob_start();

        self::$writer->save('php://output');

        ob_end_flush();
        
        exit;
    }

    private static function getColunaLetra($index)
    {
        return chr(65 + $index - 1);
    }

    private static function setColumnAndField($colunas)
    {
        foreach($colunas as $titulo){
            array_push(self::$column, $titulo[0]);
            array_push(self::$field, $titulo[1]);
         }
    }

    private static function setArrayDados($resultado, $dadosBD)
    {
        if($dadosBD){ 
            while($row = sqlsrv_fetch_array($resultado)){
                $linha = array();
                foreach($row as $coluna => $valor){
                    if(!is_numeric($coluna)){ $linha[$coluna] = (is_object($valor)) ? date_format($valor, "d/m/Y H:i") : $valor; }
                }
                array_push(self::$dados, $linha); 
            }
        }else{ self::$dados = $resultado; }
    }

    private static function setFormartDados($dado, $field)
    {
        switch($field[1]){
            case 'date':
                return explode(' ', $dado[$field[0]])[0];
                break;
            case 'time':
                return explode(' ', $dado[$field[0]])[1];
                break;
            case 'contador_auto':
                return ($field[0] == null) ? self::$contador_auto : $field[0] . self::$contador_auto;
                break;
            case 'conteudo_fixo':
                return $field[0];
                break;
            case 'somente_numeros':
                return preg_replace('/[^0-9]/', '', $dado[$field[0]]);
                break;
            default:
                return call_user_func_array($field[1], array($dado[$field[0]]));
                break;
        }
    }

    
}


?>