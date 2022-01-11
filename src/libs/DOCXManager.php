<?php

/**
 * @author ITERNOVA [https://www.iternova.net]
 * @version 2.1.7 - 20220112
 * @package iternova\docxmerger
 */

namespace Iternova\DOCXMerger\libs;

/**
 * Class DOCXManager Read & Write DOCX files
 */
class DOCXManager {

    /** @var string $docx_path Path al fichero DOCX actual */
    private $docx_path;
    /** @var bool|string $docx_rels Datos de _RELS del DOCX actual */
    private $docx_rels;
    /** @var bool|string Datos del DOCX actual */
    private $docx_document;
    /** @var string $docx_content_types Datos de ContentType del DOCX actual */
    private $docx_content_types;
    /** @var TbsZip $tbszip Fichero ZIP del DOCX actual */
    private $tbszip;
    /** @var array $array_headers_footers Array "Path del ZIP" => "Contenido" */
    private $array_headers_footers = [];

    // Constantes requeridas para generar el DOCX
    const RELS_ZIP_PATH = "word/_rels/document.xml.rels";
    const DOC_ZIP_PATH = "word/document.xml";
    const CONTENT_TYPES_PATH = "[Content_Types].xml";
    const ALT_CHUNK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
    const ALT_CHUNK_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";


    /**
     * Constructor
     * @param string $docx_path Path al fichero DOCX actual
     */
    public function __construct( $docx_path ) {
        $this->docx_path = $docx_path;

        $this->tbszip = new TbsZip();
        $this->tbszip->Open( $this->docx_path );

        $this->docx_rels = $this->read_content( self::RELS_ZIP_PATH );
        $this->docx_document = $this->read_content( self::DOC_ZIP_PATH );
        $this->docx_content_types = $this->read_content( self::CONTENT_TYPES_PATH );
    }

    /**
     * @param string $zip_file_path Ruta absoluta al fichero ZIP
     * @return bool|string Datos leidso del fichero ZIP
     */
    private function read_content( $zip_file_path ) {
        return $this->tbszip->FileRead( $zip_file_path );
    }

    /**
     * Escribe contenido en fichero DOCX
     * @return int
     */
    private function write_content( $content, $zipPath ) {
        $this->tbszip->FileReplace( $zipPath, $content, TBSZIP_STRING );
        return 0;
    }

    /**
     * Introduce fichero en DOCX
     * @param string $file_path Ruta absolta al fichero DOCX a incluir
     * @param string $zip_file_name Nombre del fichero ZIP resultante
     * @param string|int $refID Identificador de referencia
     * @param bool $page_break Introducir salto de pagina antes del DOCX
     */
    public function add_file( $file_path, $zip_file_name, $refID, $page_break = false ) {
        $content = file_get_contents( $file_path );
        $this->tbszip->FileAdd( $zip_file_name, $content );

        $this->add_reference( $zip_file_name, $refID );
        $this->add_alternative_chunk( $refID, $page_break );
        $this->add_content_type( $zip_file_name );
    }

    /**
     * Introduce referencia dentro de fichero DOCX
     * @param string $zip_file_name Nombre del fichero ZIP resultante
     * @param string|int $refID Identificador de referencia
     */
    private function add_reference( $zip_file_name, $refID ) {
        $relation_xml = '<Relationship Target="../' . $zip_file_name . '" Type="' . self::ALT_CHUNK_TYPE . '" Id="' . $refID . '"/>';
        $position = strpos( $this->docx_rels, '</Relationships>' );
        $this->docx_rels = substr_replace( $this->docx_rels, $relation_xml, $position, 0 );
    }

    /**
     * Introduce datos alternativos de bloque / chunk
     * @param string|int $refID Identificador de referencia
     * @param bool $page_break Introducir salto de pagina antes del DOCX
     */
    private function add_alternative_chunk( $refID, $page_break ) {
        $page_break_xml = $page_break ? '<w:p><w:r><w:br w:type="page" /></w:r></w:p>' : '';
        $ref_xml = $page_break_xml . '<w:altChunk r:id="' . $refID . '"/>';

        $position = strpos( $this->docx_document, '</w:body>' );
        $this->docx_document = substr_replace( $this->docx_document, $ref_xml, $position, 0 );
    }


    /**
     * Introduce ContentType del fichero incluido dentro de fichero DOCX
     * @param string $zip_file_name Nombre del fichero ZIP resultante
     */
    private function add_content_type( $zip_file_name ) {
        $contenttype_xml = '<Override ContentType="' . self::ALT_CHUNK_CONTENT_TYPE . '" PartName="/' . $zip_file_name . '"/>';

        $position = strpos( $this->docx_content_types, '</Types>' );
        $this->docx_content_types = substr_replace( $this->docx_content_types, $contenttype_xml, $position, 0 );
    }

    /**
     * Carga cabeceras y pies de pagina en fichero DOCX
     * @throws \Exception
     */
    public function load_headers_footers() {
        $xml_docx_relations = new \SimpleXMLElement( $this->docx_rels );
        foreach ( $xml_docx_relations as $rel ) {
            if ( \in_array( $rel[ "Type" ], [ "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" ], true ) ) {
                $path = "word/" . $rel[ "Target" ];
                $this->array_headers_footers[ $path ] = $this->read_content( $path );
            }
        }
    }

    /**
     * Busca y remplaza una clave dentro del fichero DOCX (contenido, headers and footers) y la reemplaza con un nuevo valor
     * @param string $key Clave a buscar
     * @param string $value Valor a reemplazar
     */
    public function find_and_replace( $key, $value ) {
        $this->docx_document = str_replace( $key, $value, $this->docx_document );
        foreach ( $this->array_headers_footers as $path => $content ) {
            $this->array_headers_footers[ $path ] = str_replace( $key, $value, $content );
        }
    }

    /**
     * Escribe en disco el fichero DOCX generado
     */
    public function flush() {
        $this->write_content( $this->docx_rels, self::RELS_ZIP_PATH );
        $this->write_content( $this->docx_document, self::DOC_ZIP_PATH );
        $this->write_content( $this->docx_content_types, self::CONTENT_TYPES_PATH );
        foreach ( $this->array_headers_footers as $path => $content ) {
            $this->write_content( $content, $path );
        }

        // Almacena en otro fichero resultante. Si se guarda en el mismo fichero el ZIP se corrompe.
        // Una vez generado, se reemplaza
        $tempFile = tempnam( dirname( $this->docx_path ), "dm" );
        $this->tbszip->Flush( TBSZIP_FILE, $tempFile );

        if ( stripos( PHP_OS, 'WIN' ) === 0 ) {
            copy( $tempFile, $this->docx_path );
            unlink( $tempFile );
        } else {
            rename( $tempFile, $this->docx_path );
        }
    }

}
