<?php
require_once '../PHPWord.php';

/**
 * @link http://msdn.microsoft.com/en-us/library/office/ee922775.aspx#odc_Office14_ta_WorkingWithNumbering_Overview
 */

// New Word Document
$PHPWord = new PHPWord();

$level1 = new PHPWord_Style_Paragraph();
$level1->setTabs(new PHPWord_Style_Tabs(array(
    new PHPWord_Style_Tab('clear', 720),
    new PHPWord_Style_Tab('num', 360)
)));
$level1->setIndentions(new PHPWord_Style_Indentation(array(
    'left' => 360,
    'hanging' => 360
)));

$level2 = new PHPWord_Style_Paragraph();
$level2->setTabs(new PHPWord_Style_Tabs(array(
    new PHPWord_Style_Tab('left', 720),
    new PHPWord_Style_Tab('num', 720)
)));
$level2->setIndentions(new PHPWord_Style_Indentation(array(
    'left' => 720,
    'hanging' => 360
)));

$bulletFont = new PHPWord_Style_Font();
$bulletFont->setName("Symbol");

$bulletFont1 = new PHPWord_Style_Font();
$bulletFont1->setName("Wingdings");

/* Examples of the two ways to create bulleted lists */
$bulleted = new PHPWord_Numbering_AbstractNumbering("Bulleted");
$bulleted->addLevel(new PHPWord_Numbering_Level("1", PHPWord_Numbering_Level::NUMFMT_BULLET, "", "left", $dummy = NULL, $bulletFont));

$numbered = new PHPWord_Numbering_AbstractNumbering("Numbered", array(
    new PHPWord_Numbering_Level("1", PHPWord_Numbering_Level::NUMFMT_DECIMAL, "%1.", "left")
));

$ordinalLeft = new PHPWord_Numbering_AbstractNumbering("Ordinal_Left", array(
    new PHPWord_Numbering_Level("1", PHPWord_Numbering_Level::NUMFMT_ORDINAL_TEXT, "%1)", "left")
));

$ordinalRight = new PHPWord_Numbering_AbstractNumbering("Ordinal_Right", array(
    new PHPWord_Numbering_Level("1", PHPWord_Numbering_Level::NUMFMT_ORDINAL_TEXT, "%1)", "right")
));

$simpleMultiLevel = new PHPWord_Numbering_AbstractNumbering("Simple Multi-level", array(
    new PHPWord_Numbering_Level("1", PHPWord_Numbering_Level::NUMFMT_DECIMAL, "%1.", "left"),
    new PHPWord_Numbering_Level("1", PHPWord_Numbering_Level::NUMFMT_LOWER_LETTER, "%2.", "left")
));

$advancedMultiLevel = new PHPWord_Numbering_AbstractNumbering("Adv Multi-level", array(
    new PHPWord_Numbering_Level("1", PHPWord_Numbering_Level::NUMFMT_DECIMAL, "%1.", "left", $level1),
    new PHPWord_Numbering_Level("1", PHPWord_Numbering_Level::NUMFMT_LOWER_LETTER, "%2.", "left", $level2)
));

$advancedMultiBullet = new PHPWord_Numbering_AbstractNumbering("Adv Multi-level Bullet", array(
    new PHPWord_Numbering_Level("1", PHPWord_Numbering_Level::NUMFMT_BULLET, "", "left", $level1, $bulletFont1),
    new PHPWord_Numbering_Level("1", PHPWord_Numbering_Level::NUMFMT_BULLET, "", "left", $level2, $bulletFont)
));

$PHPWord->addNumbering($bulleted);
$PHPWord->addNumbering($numbered);
$PHPWord->addNumbering($ordinalLeft);
$PHPWord->addNumbering($ordinalRight);
$PHPWord->addNumbering($simpleMultiLevel);
$PHPWord->addNumbering($advancedMultiLevel);
$PHPWord->addNumbering($advancedMultiBullet);

// New portrait section
$section = $PHPWord->createSection();

// Figure 3. Bulleted list items
$section->addListItem('One',   0, $bulleted);
$section->addListItem('Two',   0, $bulleted);
$section->addListItem('Three', 0, $bulleted);
$section->addTextBreak(1);

// Figure 4. Numbered list items
$section->addListItem('Paragraph one.',   0, $numbered);
$section->addListItem('Paragraph two.',   0, $numbered);
$section->addListItem('Paragraph three.', 0, $numbered);
$section->addTextBreak(1);

// Figure 6. Justified list items
$section->addListItem('This is the first paragraph.', 0, $ordinalLeft);
$section->addListItem('Here is the second.',          0, $ordinalLeft);
$section->addListItem('And finally a third.',         0, $ordinalLeft);
$section->addTextBreak(1);
$section->addListItem('This is the first paragraph.', 0, $ordinalRight);
$section->addListItem('Here is the second.',          0, $ordinalRight);
$section->addListItem('And finally a third.',         0, $ordinalRight);
$section->addTextBreak(1);

// Figure 9. A simple multi-level list
$section->addListItem('One', 0, $simpleMultiLevel);
$section->addListItem('Two', 1, $simpleMultiLevel);
$section->addTextBreak(1);

// My "advanced" numbered example
$PHPWord->addFontStyle('myOwnStyle', array('color'=>'FF0000'));
$section->addListItem('Lorem', 0, $advancedMultiLevel, 'myOwnStyle');
$section->addListItem('Nullam: tristique sollicitudin mattis. Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 1, $advancedMultiLevel);
$section->addListItem('Lorem: ipsum dolor sit amet, consectetur adipiscing elit. Nullam convallis nunc sit amet nulla consectetur ac fermentum nunc semper. In commodo.', 1, $advancedMultiLevel);
$section->addListItem('Commodo: orci odio. Nunc lectus purus, mollis ac euismod quis, molestie eu.', 1, $advancedMultiLevel);
$section->addListItem('Suspendisse: condimentum vulputate venenatis. Quisque placerat consectetur eleifend. Duis pulvinar odio risus. Aliquam in sapien turpis, sed varius dolor. Mauris dignissim.', 1, $advancedMultiLevel);
$section->addListItem('Ipsum', 0, $advancedMultiLevel);
$section->addListItem('Pellentesque: feugiat laoreet elit, ac viverra quam dignissim ac.', 1, $advancedMultiLevel);
$section->addListItem('Cras: tincidunt accumsan dolor in faucibus. Quisque est nisl, porta a sollicitudin vel, tincidunt at odio. Aliquam.', 1, $advancedMultiLevel);
$section->addTextBreak(1);

// My "advanced" bullet example
$section->addListItem('Lorem', 0, $advancedMultiBullet);
$section->addListItem('Nullam: tristique sollicitudin mattis. Lorem ipsum dolor sit amet, consectetur adipiscing elit.', 1, $advancedMultiBullet);
$section->addListItem('Lorem: ipsum dolor sit amet, consectetur adipiscing elit. Nullam convallis nunc sit amet nulla consectetur ac fermentum nunc semper. In commodo.', 1, $advancedMultiBullet);
$section->addListItem('Commodo: orci odio. Nunc lectus purus, mollis ac euismod quis, molestie eu.', 1, $advancedMultiBullet);
$section->addListItem('Suspendisse: condimentum vulputate venenatis. Quisque placerat consectetur eleifend. Duis pulvinar odio risus. Aliquam in sapien turpis, sed varius dolor. Mauris dignissim.', 1, $advancedMultiBullet);
$section->addListItem('Ipsum', 0, $advancedMultiBullet);
$section->addListItem('Pellentesque: feugiat laoreet elit, ac viverra quam dignissim ac.', 1, $advancedMultiBullet);
$section->addListItem('Cras: tincidunt accumsan dolor in faucibus. Quisque est nisl, porta a sollicitudin vel, tincidunt at odio. Aliquam.', 1, $advancedMultiBullet);
$section->addTextBreak(1);

// Save File
$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
$objWriter->save('CustomListItem.docx');
?>