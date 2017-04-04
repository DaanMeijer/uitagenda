<?PHP

namespace Merel\Agenda;

class Parser implements Listener {
	
	private $audience = 'main';
	private $month = '';
	private $day_of_month = 0;
	private $day_of_week = '';
	private $category = '';
	
	private $stack = [];
	public function push($text){
		$this->stack[] = $text;
	}
	
	private $listeners = [];
	
	public function addListener(Listener $listener){
		$this->listeners[] = $listener;
	}
	
	public function haveMonth($month){
		$this->month = $month;
		foreach($this->listeners as $listener){
			$listener->haveMonth($month);
		}
	}
	
	public function haveDay($day, $matches = []){

		foreach($this->listeners as $listener){
			$listener->haveDay($day);
		}

		if(count($matches) > 1) {
			$this->day_of_month = $matches[1];
		}

		if(count($matches) > 2) {
			$this->day_of_week = $matches[2];
		}

		$this->category = '';
	}
	
	public function haveCategory($category){
		foreach($this->listeners as $listener){
			$listener->haveCategory($category);
		}
		$this->category = $category;
	}
	
	public function haveLocation($location){
		foreach($this->listeners as $listener){
			$listener->haveLocation($location);
		}
	}
	
	public function haveTitle($title){
		foreach($this->listeners as $listener){
			$listener->haveTitle($title);
		}
	}

	public function haveDescription($description){
		foreach($this->listeners as $listener){
			$listener->haveDescription($description);
		}
	}

	public function haveAudience($audience){
		foreach($this->listeners as $listener){
			$listener->haveAudience($audience);
		}
	}

	public function havePremiere($premiere){
		foreach($this->listeners as $listener){
			$listener->havePremiere($premiere);
		}
	}
	
	public function noMatch($combinedItem){
		foreach($this->listeners as $listener){
			$listener->noMatch($combinedItem);
		}
	}
	
	public function haveCombinedItem($combinedItem){
		
		if(preg_match_all('/([^>.]+?) *> ?([^\\]]+\\])? ?([^>]+) */', $combinedItem, $matches, PREG_SET_ORDER)){
			foreach($matches as $match){
				$this->haveLocation($match[1] . ' >');
				if(count($match) > 3){
					$this->haveTitle($match[2]);
					$this->haveDescription($match[3]);
				}else{
					$this->haveDescription($match[2]);
				}
			}
		}else{
			$this->noMatch($combinedItem);
		}
		
	}
	
	public function textBreak(){
		switch(count($this->stack)){
			case 0:
				return;
				
			case 1:
				$text = $this->stack[0];
				if(preg_match('/^([0-9]+) +([a-zA-Z]+day) *$/', $text, $matches)){
					$this->haveDay($text, $matches);
				}elseif(in_array($text, [
					'January',
					'February',
					'March',
					'April',
					'May',
					'June',
					'July',
					'August',
					'September',
					'October',
					'November',
					'December',
					'Movies',
					'Festivals & Events',
					'Exhibitions & Museums',
					'Kids',
				])){
					$this->haveMonth($text);
				}elseif(strlen($text) > 64){
					$this->haveCombinedItem($text);
				}else{
					$this->haveCategory($text);
				}
				break;
				
			default:
				//item!
				$text = implode('', $this->stack);
				
				if(strlen($text) > 64){
					$this->haveCombinedItem($text);
				}else{
					$this->haveCategory($text);
				}
				

		}
		
		//printf("Stack (%02d): %s\n", count($this->stack), implode('', $this->stack));
		
		$this->stack = [];
	}
	
	public function parse($fileName, $mimeType){
		
		switch($mimeType){
			case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
				$this->parseDocx($fileName);
				break;

			case null:
			case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
				$this->parseXlsx($fileName);
				break;
			default:
				var_dump($mimeType, $fileName);
				die();
				break;
		}
		
	}
	
	public function parseXlsx($fileName){
		$type = \PHPExcel_IOFactory::identify($fileName);
		$reader = \PHPExcel_IOFactory::createReader($type);
		$reader->setReadDataOnly(TRUE);

		$phpExcel = $reader->load($fileName);

		foreach($phpExcel->getAllSheets() as $sheet){

			$lookup = [];
			$col = 0;
			do {
				$name = $sheet->getCellByColumnAndRow($col, 1)->getValue();
				if($name){
					$lookup[strtolower($name)] = $col;
				}
				$col++;
			}while($name);

			if(!isset($lookup['locatienaam'])){
				continue;
			}


			$this->haveAudience($sheet->getTitle());


//			var_dump($sheet->getTitle(), $lookup);

			$highestRow = $sheet->getHighestRow();

			for($a=2; $a<$highestRow; $a++){





				switch(strtolower($sheet->getTitle())){
					case 'agenda':
					case 'leidsche rijn':
					case 'kids':

						$location = $sheet->getCellByColumnAndRow($lookup['locatienaam'], $a)->getValue();

						if(!$location){
							continue;
						}

						$dayCell = $sheet->getCellByColumnAndRow($lookup['van'], $a);
						$day = $dayCell->getFormattedValue();

						if($day){
							$date =\PHPExcel_Shared_Date::ExcelToPHP($day);
							$timestamp = new \DateTime();
							$timestamp->setTimestamp($date);

							$day = $timestamp->format("j l");
							$matches = [$day, $timestamp->format("j"), $timestamp->format("l")];

							$this->haveDay($day, $matches);
						}

						$category = $sheet->getCellByColumnAndRow($lookup['type'], $a)->getValue();
						if($category) {
							$this->haveCategory( $category );
						}

						$this->haveLocation($location . ' >');

						$title = $sheet->getCellByColumnAndRow($lookup['titel'], $a)->getValue();

						$time = $sheet->getCellByColumnAndRow($lookup['tijdstip'], $a)->getFormattedValue();


						if(is_numeric($time)){
							$time = $time * (24*60);
							$time = sprintf('[%02d:%02d]', floor($time / 60), $time % 60);
						}

						if($time){
							$title .= ' ' . $time;
						}

						$this->haveTitle($title);


						$description = $sheet->getCellByColumnAndRow($lookup['korte beschrijving'], $a)->getValue();
						$description = preg_replace('/_x000D_/', '', $description);
						$this->haveDescription($description);

						break;

					case 'film':

						$dayCell = $sheet->getCellByColumnAndRow($lookup['van'], $a);
						$day = $dayCell->getFormattedValue();

						if(!$day){
							continue;
						}

						$location = $sheet->getCellByColumnAndRow($lookup['locatienaam'], $a)->getValue();
						if($location){
							$this->haveDay($location);
						}

						$this->haveLocation($day);

						$premiere = $sheet->getCellByColumnAndRow($lookup['première'], $a);
						if($premiere){
							$this->havePremiere($premiere);
						}

						$title = $sheet->getCellByColumnAndRow($lookup['titel'], $a)->getValue();

						$time = $sheet->getCellByColumnAndRow($lookup['tijdstip'], $a)->getFormattedValue();


						if(is_numeric($time)){
							$time = $time * (24*60);
							$time = sprintf('[%02d:%02d]', floor($time / 60), $time % 60);
						}

						if($time){
							$title .= ' ' . $time;
						}

						$this->haveTitle($title);

						$description = $sheet->getCellByColumnAndRow($lookup['korte beschrijving'], $a)->getValue();
						$description = preg_replace('/_x000D_/', '', $description);
						$this->haveDescription($description);

						break;

					case 'musea & expo':



						$dayCell = $sheet->getCellByColumnAndRow($lookup['van'], $a);
						$day = $dayCell->getFormattedValue();

						if(!$day){
							continue;
						}

						$location = $sheet->getCellByColumnAndRow($lookup['locatienaam'], $a)->getValue();
						if($location){
							$this->haveCategory($location);
						}


						$this->haveLocation($day . ' >');

						$title = $sheet->getCellByColumnAndRow($lookup['titel'], $a)->getValue();
						$this->haveTitle($title);

						$description = $sheet->getCellByColumnAndRow($lookup['korte beschrijving'], $a)->getValue();
						$description = preg_replace('/_x000D_/', '', $description);
						$this->haveDescription($description);

						break;

					case 'festival & evenementen':

						$location = $sheet->getCellByColumnAndRow($lookup['locatienaam'], $a)->getValue();

						if(!$location){
							continue;
						}

						$title = $sheet->getCellByColumnAndRow($lookup['titel - datum'], $a)->getValue();
						$this->haveCategory($title);

						if($location){
							$this->haveLocation($location . ' >');
						}

						$time = $sheet->getCellByColumnAndRow($lookup['tijdstip'], $a)->getFormattedValue();

						if(is_numeric($time)){
							$time = $time * (24*60);
							$time = sprintf('[%02d:%02d]', floor($time / 60), $time % 60);
						}
						$this->haveTitle($time);

						$description = $sheet->getCellByColumnAndRow($lookup['korte beschrijving'], $a)->getValue();
						$description = preg_replace('/_x000D_/', '', $description);
						$this->haveDescription($description);

						break;



					default:
						var_dump($sheet->getTitle());
						die();
						break;
				}

			}
		}

	}
	
	public function parseDocx($fileName){
		$phpWord = \PhpOffice\PhpWord\IOFactory::load($fileName);
	
		$sections = $phpWord->getSections();
		foreach($sections as $section){
			$elements = $section->getElements();
			foreach($elements as $element){
				$class = get_class($element);
				switch($class){
					case 'PhpOffice\PhpWord\Element\Text':
						//echo "Text: " . $element->getText();
						$this->textBreak();
						$this->push($element->getText());
						$this->textBreak();
						break;
						
					case 'PhpOffice\PhpWord\Element\TextRun':
						$text = '';
						foreach($element->getElements() as $subElement){
							$subclass = get_class($subElement);
							switch($subclass){
								case 'PhpOffice\PhpWord\Element\Text':
									$text .= $subElement->getText();
									break;
									
								default:
									echo "Unknown subelement: {$subclass}";
							}
						}
						$this->push($text);
						//echo "Text: $text";
						break;
						
					case 'PhpOffice\PhpWord\Element\TextBreak':
						$this->textBreak();
						break;
					default:
						echo "Unknown class: {$class}";
						echo "\n";
				}

			}
		}
	}
	
}
