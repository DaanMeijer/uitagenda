<?PHP
namespace Merel\Agenda;

interface Listener {
	public function haveMonth($month);
	public function haveDay($day);
	public function haveCategory($category);
	public function haveLocation($location);
	public function haveTitle($title);
	public function haveDescription($description);
	
	public function noMatch($combinedItem);
}