#!/usr/bin/php

<?php
include('gChart.php');
include('BulletGraph.php');

$plannerScore = $argv[1];
$licenseeScore = $argv[2];
$industryScore = $argv[3];
$output_file = $argv[4];


// $bulletGraph = new BulletGraph(880, 60, "Your Score");
$bulletGraph = new BulletGraph(575, 50, " ");

$bulletGraph->setBrokers($licenseeScore);
$bulletGraph->setMarker($industryScore);
$bulletGraph->setYourscore($plannerScore);

//$bulletGraph->setBrokers($params["moc-average"]);
//$bulletGraph->setMarker($params["brokers-average"]);
//$bulletGraph->setRange($config->getBulletGraphRanges(false));
$bulletGraph->setRange(array(60,80,100));
//$bulletGraph->setYourscore($params["your-score"]);

$bulletGraph->createGraph();

$url = html_entity_decode($bulletGraph->getImageURL());

$image = file_get_contents($url);

file_put_contents($output_file, $image);

echo "wget \"" . $url . "\" -O \"" . $output_file . "\" &"
?>
