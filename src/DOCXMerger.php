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


    /**
     * @param string $file_path Ruta absoluta de fichero DOCX a mergear
     */
    public function add_file( $file_path ) {
        $this->array_files[] = $file_path;
    }

    /**
     * @param array $array_files_paths Rutas absolutas de ficheros DOCX a mergear
     */
    public function add_files( array $array_files_paths ) {
        $this->array_files = array_merge( $this->array_files, $array_files_paths );
    }


    /**
     * @param string $output_file_path Ruta absoluta al fichero DOCX resultante del mergeo
     * @param bool $page_breaks Incluir saltos de pagina tras cada fichero
     * @return bool Resultado operacion
     * @throws \RuntimeException
     */
    public function save( $output_file_path, $page_breaks = false ) {
        if ( !count( $this->array_files ) ) {
            return false;
        }

        if ( !copy( $this->array_files[ 0 ], $output_file_path ) ) {
            throw new \RuntimeException( "Error saving output file: " . $output_file_path );
        }

        $obj_file_docx = new \Iternova\DOCXMerger\libs\DOCXManager( $output_file_path );
        foreach ( $this->array_files as $file_index => $file_path ) {
            $obj_file_docx->add_file( $file_path, "file_part_" . $file_index . ".docx", "rId10" . $file_index, $page_breaks );
        }

        $obj_file_docx->flush();
        return true;
    }

}