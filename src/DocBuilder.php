<?PHP
namespace Merel\Agenda;

use \PhpOffice\PhpWord\PhpWord;
use \PhpOffice\PhpWord\Settings;

class DocBuilder implements Listener {
	
	private $doc = null;
	public function __construct(){
		
		Settings::setOutputEscapingEnabled(true);
		
		$this->doc = new PHPWord();
		$this->section = $this->doc->addSection();
		
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
				
		$this->doc->addFontStyle('fsLocation', [
			'name' => 'Tahoma',
			'size' => 12,
			'color' => '000000',
			'bold' => true,
			'allCaps' => true,
			'underline' => 'solid',
		]);		
		
		$this->doc->addFontStyle('fsTitle', [
			'name' => 'Tahoma',
			'size' => 12,
			'color' => '000000',
		]);		
		
		$this->doc->addFontStyle('fsDescription', [
			'name' => 'Tahoma',
			'size' => 12,
			'color' => '000000',
			'italic' => true,
		]);
	}
	
	public function haveMonth($month){
		$this->section->addTitle($month, 1);
	}
	
	public function haveDay($day){
		$this->section->addTitle($day, 2);
	}
	
	public function haveCategory($category){
		$this->section->addTitle($category, 3);
	}
	
	
	
	public function haveLocation($location){
		$this->textRun = $this->section->addTextRun();
		
		$textElm = $this->textRun->addText($location);
		$textElm->setFontStyle('fsLocation');
	}
	
	public function haveTitle($title){
		$textElm = $this->textRun->addText(' ' . $title);
		$textElm->setFontStyle('fsTitle');
	}
	
	public function haveDescription($description){
		$textElm = $this->textRun->addText(' ' . $description);
		$textElm->setFontStyle('fsDescription');
	}
	
	public function render(){
		// save as a random file in temp file
		$temp_file = tempnam(sys_get_temp_dir(), 'PHPWord');
		$this->doc->save($temp_file);

		// Your browser will name the file "myFile.docx"
		// regardless of what it's named on the server 
		header("Content-Disposition: attachment; filename='agenda.docx'");
		readfile($temp_file); // or echo file_get_contents($temp_file);
		unlink($temp_file);  // remove temp file
	}
	
	public function noMatch($combinedItem){
		
	}
	
}
