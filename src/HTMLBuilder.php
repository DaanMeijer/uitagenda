<?PHP
namespace Merel\Agenda;

class HTMLBuilder implements Listener {
	
	private $parts = [];
	
	public function haveMonth($month){
		$this->add(sprintf("<span class='month'>%s</span>\n", htmlentities($month)));
	}
	
	public function haveDay($day){
		$this->add(sprintf("<span class='day'>%s</span>\n", htmlentities($day)));
	}
	
	public function haveCategory($category){
		$this->add(sprintf("<span class='category'>%s</span>\n", htmlentities($category)));
	}
	
	public function haveLocation($location){
		$this->add(sprintf("<span class='item'><span class='location'>%s</span>\n", htmlentities($location)));
	}
	
	public function haveTitle($title){
		$this->add(sprintf("<span class='title'>%s</span>\n", htmlentities($title)));
	}
	
	public function haveDescription($description){
		$this->add(sprintf("<span class='description'>%s</span></span>\n", htmlentities($description)));
	}
	
	private function style(){
		?>
		<style>
		.month {
			display: block;
			width: 200px;
			height: 200px;
			background: red;
			display: block;
		}
		
		.day {
			display: block;
			font-weight: bold;
			font-size: 20px;
		}
		
		.category {
			display: block;
			color: red;
			font-style: italic;
			margin: 0;
		}
		
		.location, .title, .description {
			display: inline;
			font-size: 12px;
			margin: 0;
		}
		
		body .location {
			font-weight: bold;
			display: inline-block;
			text-transform: uppercase;
			text-decoration: underline;
		}
		
		.description {
			font-style: italic;
		}
		
		.item {
			display: block;
		}
		</style>
		<?PHP
	}
	
	public function render(){
		$this->style();
		
		echo implode('', $this->parts);
	}
	
	private function add($item){
		$this->parts[] = $item;
	}
	
	public function noMatch($combinedItem){
		$len = strlen($combinedItem);
		echo "<p style='color: red'>No match for |{$combinedItem}| ({$len})</p>";
		if(@$_GET['_debug']){
			echo "<pre>";
			debug_print_backtrace();
			echo "</pre>";
		}
	}
	
}