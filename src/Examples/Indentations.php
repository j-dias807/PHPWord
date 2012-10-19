<?php
require_once '../PHPWord.php';

// New Word Document
$PHPWord = new PHPWord();

// New portrait section
$section = $PHPWord->createSection();

// Create a bunch of Styles
$numberedList = array(
  'listType' => PHPWord_Style_ListItem::TYPE_NUMBER,
  'indentations' => array(
      'left' => 750
  )
);

$PHPWord->addParagraphStyle('indented', array(
  'spaceAfter' => 95,
  'indentations' => array(
    'left' => 750,
    'right' => 750
  )
));

// Add listitem elements
$section->addListItem('List Item 1', 0, null, $numberedList);
$section->addListItem('List Item 2', 0, null, $numberedList);
$section->addListItem('List Item 3', 0, null, $numberedList);
$section->addTextBreak(2);

// Add text section
$section->addText("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent et lorem lorem. Quisque lacinia metus a odio cursus vitae ultrices est porttitor. Donec sit amet dignissim dolor. In eu diam volutpat nulla varius ullamcorper in eu nunc. Nullam pretium faucibus condimentum. In eget eros tortor, eget congue augue. Phasellus vestibulum, erat nec placerat ullamcorper, quam velit luctus mi, et volutpat mauris nunc vitae magna. Curabitur augue nulla, lacinia a volutpat vitae, commodo vitae erat.", NULL, "indented");
$section->addTextBreak(2);

// Save File
$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
$objWriter->save('Indentations.docx');
?>
