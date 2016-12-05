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
		$this->day_of_month = $matches[1];
		$this->day_of_week = $matches[2];
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
				])){
					$this->haveMonth($text);
				}elseif(strlen($text) > 50){
					$this->haveCombinedItem($text);
				}else{
					$this->haveCategory($text);
				}
				break;
				
			default:
				//item!
				$text = implode('', $this->stack);
				
				if(strlen($text) > 50){
					$this->haveCombinedItem($text);
				}else{
					$this->haveCategory($text);
				}
				

		}
		
		//printf("Stack (%02d): %s\n", count($this->stack), implode('', $this->stack));
		
		$this->stack = [];
	}
	
	public function parse($fileName){
		
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