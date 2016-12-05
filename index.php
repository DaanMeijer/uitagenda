<?PHP


if($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['wordfile'])){
	require_once(__DIR__ . '/vendor/autoload.php');
	

	$fileName = $_FILES['wordfile']['tmp_name'];
	
	$parser = new Merel\Agenda\Parser();
	
	/** /
	$listener = new Merel\Agenda\HTMLBuilder();
	/*/
	$listener = new Merel\Agenda\DocBuilder();
	/**/
	$parser->addListener($listener);
	
	$parser->parse($fileName);
	
	$listener->render();
	?><?PHP
	
}else{
	?>
<!DOCTYPE html>
<html>
<body>

<form action="" method="post" enctype="multipart/form-data">
    Select file to upload:
    <input type="file" name="wordfile" id="fileToUpload">
    <input type="submit" value="Upload File" name="submit">
</form>

</body>
</html>
	<?PHP
}