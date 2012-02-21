<?php

class Word extends ZipArchive{

    // Файлы для включения в архив
    private $files;

    // Путь к шаблону
    public $path;

    // Содержимое документа
    protected $content;

    public function __construct($filename, $template_path = '/template/' ){

      // Путь к шаблону
      $this->path = dirname(__FILE__) . $template_path;

      // Если не получилось открыть файл, то жизнь бессмысленна.
      if ($this->open($filename, ZIPARCHIVE::CREATE) !== TRUE) {
        die("Unable to open <$filename>\n");
      }


      // Структура документа
      $this->files = array(
        "word/_rels/document.xml.rels",
        "word/theme/theme1.xml",
        "word/fontTable.xml",
        "word/settings.xml",
        "word/styles.xml",
        "word/stylesWithEffects.xml",
        "word/webSettings.xml",
        "_rels/.rels",
        "docProps/app.xml",
        "docProps/core.xml",
        "[Content_Types].xml" );

      // Добавляем каждый файл в цикле
      foreach( $this->files as $f )
        $this->addFile($this->path . $f , $f );
    }

    // Регистрируем текст
    public function assign( $text = '' ){

      // Берем шаблон абзаца
      $p = file_get_contents( $this->path . 'p.xml' );

      // Нам нужно разбить текст по строкам
      $text_array = explode( "\n", $text );

      foreach( $text_array as $str )
        $this->content .= str_replace( '{TEXT}', $str, $p );
    }

    // Упаковываем архив
    public function create(){

      // Добавляем содержимое
      $this->addFromString("word/document.xml", str_replace( '{CONTENT}', $this->content, file_get_contents( $this->path . "word/document.xml" ) ) );

      $this->close();
    }
}
