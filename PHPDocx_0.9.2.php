<?php

/**************************************
/
/  *** PHP Library PHPDocx ***
/  Version: 0.9.2
/  Author: Alexey Kichaev
/  Home page: http://webli.ru/phpdocx/
/  License: MIT or GPL
/  GITHub: https://github.com/alkich/PHPDocx
/  Date: 14/03/2012
/
/***************************************/

// Общий класс для создания генераторов MS Office документов
class OfficeDocument extends ZipArchive{

    // Путь к шаблону
    protected $path;

    // Содержимое документа
    protected $content;

    // Множитель для перевода размеров изображений из пикселей в EMU
    protected $px_emu = 8625;

    // Делаем приватно, чтобы не было возможности вшить дрянь в документ
    protected $rels = array();

    public function __construct($filename, $template_path = '/template/' ){

      // Путь к шаблону
      $this->path = dirname(__FILE__) . $template_path;

      // Если не получилось открыть файл, то жизнь бессмысленна.
      if ( $this->open( $filename, ZIPARCHIVE::CREATE) !== TRUE) {
        die("Unable to open <$filename>\n");
      }

      // Описываем связи для документа MS Office
      $this->rels = array_merge( $this->rels, array(
        'rId3' => array(
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
          'docProps/app.xml' ),
        'rId2' => array(
          'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
          'docProps/core.xml' ),
      ) );

      // Добавляем типы данных
      $this->addFile($this->path . "[Content_Types].xml" , "[Content_Types].xml" );
    }

    // Генерация зависимостей
    protected function add_rels( $filename, $rels, $path = '' ){

      // Шапка XML
      $xmlstring = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

      // Добавляем документы по описанным связям
      foreach( $rels as $rId => $params ){

        // Если указан путь к файлу, берем. Если нет, то берем из репозитория
        $pathfile = empty( $params[2] ) ? $this->path . $path . $params[1] : $params[2];

        // Добавляем документ в архив
        if( $this->addFile( $pathfile ,  $path . $params[1] ) === false )
          die('Не удалось добавить в архив ' . $path . $params[1] );

        // Прописываем в связях
        $xmlstring .= '<Relationship Id="' . $rId . '" Type="' . $params[0] . '" Target="' . $params[1] . '"/>';
      }

      $xmlstring .= '</Relationships>';

      // Добавляем в архив
      $this->addFromString( $path . $filename, $xmlstring );
    }

    protected function pparse( $replace, $content ){

      return str_replace( array_keys( $replace ), array_values( $replace ), $content );
    }
}

// Класс для создания документов MS Word
class WordDocument extends OfficeDocument{

    public function __construct( $filename, $template_path = '/template/' ){

      parent::__construct( $filename, $template_path );

      // Описываем связи для Word
      $this->word_rels = array(
        "rId1" => array(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
          "styles.xml"
        ),
        "rId2" => array(
          "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects",
          "stylesWithEffects.xml",
        ),
        "rId3" => array(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
          "settings.xml",
        ),
        "rId4" => array(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
          "webSettings.xml",
        ),
        "rId6" => array(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
          "fontTable.xml",
        ),
        "rId7" => array(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
          "theme/theme1.xml",
        ),
      );
    }

    public function assign( $content = '', $return = false ){

      // Проверяем, является ли $text файлом. Если да, то подключаем изображение
      if( is_file( $content ) ){

        // Берем шаблон абзаца
        $block = file_get_contents( $this->path . 'image.xml' );

        list( $width, $height ) = getimagesize( $content );

        $rid = "rId" . count( $this->word_rels ) . 'i';
        $this->word_rels[$rid] = array(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          "media/" . $content,

          // Указываем непосредственно путь к файлу
          $content
        );

        $xml = $this->pparse( array(
          '{WIDTH}' => $width * $this->px_emu,
          '{HEIGHT}' => $height * $this->px_emu,
          '{RID}' => $rid,
        ), $block );
      }
      else{

        // Берем шаблон абзаца
        $block = file_get_contents( $this->path . 'p.xml' );

        $xml = $this->pparse( array(
          '{TEXT}' => $content,
        ), $block );
      }

      // Если нам указали, что нужно возвратить XML, возвращаем
      if( $return )
        return $xml;
      else
        $this->content .= $xml;
    }

    // Упаковываем архив
    public function create(){

      $this->rels['rId1'] = array(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'word/document.xml' );

      // Добавляем связанные документы MS Office
      $this->add_rels( "_rels/.rels", $this->rels );

      // Добавляем связанные документы MS Office Word
      $this->add_rels( "_rels/document.xml.rels", $this->word_rels, 'word/' );

      // Добавляем содержимое
      $this->addFromString("word/document.xml", str_replace( '{CONTENT}', $this->content, file_get_contents( $this->path . "word/document.xml" ) ) );

      $this->close();
    }
}