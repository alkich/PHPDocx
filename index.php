<?

/**************************************
/
/  *** Test file for PHP Library PHPDocx ***
/  Version: 0.9.2
/  Author: Alexey Kichaev
/  Home page: http://webli.ru/phpdocx/
/  License: MIT or GPL
/  GITHub: https://github.com/alkich/PHPDocx
/  Date: 14/03/2012
/
/***************************************/

// Выявляем все ошибки
error_reporting( E_ALL | E_NOTICE );

// Подключаем класс
include 'PHPDocx_0.9.2.php';

// Создаем и пишем в файл. Деструктор закрывает
$w = new WordDocument( "example.docx" );

// Использование метода assign
/******************************
/
/ $w->assign( 'text' );
/ $w->assign( 'image.png' );
/ $xml = $w->assign( 'image.png', true );
/ $w->assign( $w->assign( 'image.png' ) );
/
/******************************/

$w->assign('image.jpg');
$w->assign('Кто узнал эту женщину - тот настоящий знаток женской красоты.');

$w->create();

?>