<?php

/**
 * @author ITERNOVA [https://www.iternova.net]
 * @version 2.1.7 - 20220112
 * @package iternova\docxmerger
 */

namespace Iternova\DOCXMerger;

/**
 * Class DOCXMerger Une varios ficheros DOCX
 * @note Si se utiliza para luego usarlo como plantilla de OpenTBS, es recomendable que
 *       el primer DOCX contenga los headers & footers, y el resto no tenga ninguno configurado
 */
class DOCXMerger {
    /** @var array $array_files Array de ficheros */
    private $array_files = [];

    /** @var bool[] $array_files_page_breaks Array que indica si al incluir cada fichero hay que introducir un salto de pagina previo */
    private $array_files_page_breaks = [];


    /**
     * @param string $file_path Ruta absoluta de fichero DOCX a mergear
     * @param bool $page_break Indica si se incluye un salto de pagina antes de incluir el fichero DOCX
     */
    public function add_file( $file_path, $page_break = false ) {
        $this->array_files[] = $file_path;
        $this->array_files_page_breaks[] = $page_break;
    }

    /**
     * Incluye ficheros DOCX a mergear
     * @param array $array_files_paths Rutas absolutas de ficheros DOCX a mergear
     * @param bool[] $array_files_page_breaks Indica si se incluye un salto de pagina antes de incluir el fichero DOCX correspondiente
     */
    public function add_files( $array_files_paths, $array_files_page_breaks = [] ) {
        $this->array_files = array_merge( $this->array_files, $array_files_paths );
        $this->array_files_page_breaks = array_merge( $this->array_files_page_breaks, $array_files_page_breaks );
    }


    /**
     * @param string $output_file_path Ruta absoluta al fichero DOCX resultante del mergeo
     * @return bool Resultado operacion
     * @throws \RuntimeException
     */
    public function save( $output_file_path ) {
        if ( empty( $this->array_files ) ) {
            return false;
        }

        $first_file = array_shift( $this->array_files );
        if ( !copy( $first_file, $output_file_path ) ) {
            throw new \RuntimeException( "Error saving output file: " . $output_file_path );
        }

        $obj_file_docx = new \Iternova\DOCXMerger\libs\DOCXManager( $output_file_path );
        foreach ( $this->array_files as $file_index => $file_path ) {
            $page_break = isset( $this->array_files_page_breaks[ $file_index ] ) && (bool) $this->array_files_page_breaks[ $file_index ];
            $obj_file_docx->add_file( $file_path, "file_part_" . $file_index . ".docx", "rId10" . $file_index, $page_break );
        }

        $obj_file_docx->flush();
        return true;
    }

}