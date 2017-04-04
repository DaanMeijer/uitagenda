<?PHP
namespace Merel\Agenda;

use \PhpOffice\PhpWord\PhpWord;
use \PhpOffice\PhpWord\Settings;

class DocBuilder implements Listener {

	private $doc = null;

	private $section = null;

	public function __construct() {

		Settings::setOutputEscapingEnabled( true );

		$this->doc = new PHPWord();

		$this->doc->getCompatibility()->setOoxmlVersion( 15 );


		$this->section = $this->doc->addSection();

		/*
		$this->doc->addTitleStyle(1, [
			'name' => 'Tahoma',
			'size' => 40,
			'color' => '000000',
			'bgColor' => 'ff0000',
		]);
		
		$this->doc->addTitleStyle(2, [
			'name' => 'Tahoma',
			'size' => 18,
			'color' => '000000',
			'bold' => true,
		]);
		
		$this->doc->addTitleStyle(3, [
			'name' => 'Tahoma',
			'size' => 14,
			'color' => 'ff0000',
			'italic' => true,
		]);
		*/


		$this->doc->addParagraphStyle( 'psAudience', [] );
		$this->doc->addParagraphStyle( 'psMonth', [] );
		$this->doc->addParagraphStyle( 'psDay', [] );
		$this->doc->addParagraphStyle( 'psGenre', [] );
		$this->doc->addParagraphStyle( 'psTextAgenda', [] );

		$this->doc->addFontStyle( '[none]', [
			'color' => '000000',
		] );

		$this->doc->addFontStyle( 'FsLocation', [
			'color' => '333333',
			/*
			'name' => 'Tahoma',
			'size' => 12,
			'color' => '000000',
			'bold' => true,
			'allCaps' => true,
			'underline' => 'solid',
			*/
		] );

		$this->doc->addFontStyle( 'FsTitle', [
			'color' => '666666',
			/*
			'name' => 'Tahoma',
			'size' => 12,
			'color' => '000000',
			*/
		] );

		$this->doc->addFontStyle( 'FsDescription', [
			'color' => '999999',
			/*
			'name' => 'Tahoma',
			'size' => 12,
			'color' => '000000',
			'italic' => true,
			*/
		] );

		$this->doc->addFontStyle( 'FsPremiere', [
			'color' => 'FF0000',
			/*
			'name' => 'Tahoma',
			'size' => 12,
			'color' => '000000',
			'italic' => true,
			*/
		] );

	}

	public function haveMonth( $month ) {
		$this->section->addText( $month, '[none]', 'psMonth' );
	}

	public function haveDay( $day ) {
		$this->section->addText( '', '[none]', 'psDay' );
		$this->section->addText( $day, '[none]', 'psDay' );
	}

	public function haveCategory( $category ) {
		$this->section->addText( $category, '[none]', 'psGenre' );
	}

	private $textRun = null;
	public function haveLocation( $location ) {
		$this->textRun = $this->section->addTextRun( 'psTextAgenda' );

		$textElm = $this->textRun->addText( $location, 'FsLocation', 'psTextAgenda' );
		//$textElm->setFontStyle('fsLocation');
	}

	public function haveTitle( $title ) {
		$textElm = $this->textRun->addText( ' ' . $title, 'FsTitle', 'psTextAgenda' );
		//$textElm->setFontStyle('fsTitle');
	}

	public function haveDescription( $description ) {
		$textElm = $this->textRun->addText( ' ' . $description, 'FsDescription', 'psTextAgenda' );
		//$textElm->setFontStyle('fsDescription');
	}

	public function havePremiere( $premiere ) {
		$this->textRun->addText( ' ' . $premiere, 'FsPremiere', 'psTextAgenda' );
	}

	public function haveAudience( $audience ) {
		$this->section->addText( $audience, '[none]', 'psAudience' );
	}


	public function render() {
		if ( $this->errors ) {
			die();
		}
		// save as a random file in temp file
		$temp_file = tempnam( sys_get_temp_dir(), 'PHPWord' );
		$this->doc->save( $temp_file );

		// Your browser will name the file "myFile.docx"
		// regardless of what it's named on the server 
		header( "Content-Disposition: attachment; filename='agenda.docx'" );
		readfile( $temp_file ); // or echo file_get_contents($temp_file);
		unlink( $temp_file );  // remove temp file
	}

	private $errors = false;

	public function noMatch( $combinedItem ) {
		echo "No match for: ";
		var_dump( $combinedItem );
		$this->errors = true;

	}
}