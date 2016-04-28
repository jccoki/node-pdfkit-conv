<?php
/**
 * @author David Namenyi <david@dividebyzero.com.au>
 */

class BulletGraph {

  private $brokers, $yourscore, $marker, $range;
  private $width, $height, $title;
  private $range1, $range2, $max;
  private $bulletGraph, $imageURL;

  function __construct($width, $height, $title) {
    $this->width = $width;
    $this->height = $height;
    $this->title = $title;
  }

  public function createGraph() {
  $this->bulletGraph = new gStackedBarChart($this->width, $this->height);

  $this->bulletGraph->addDataSet(array($this->yourscore));
  // creates overextended range
  if( $this->yourscore == -1 ){
    $this->bulletGraph->addDataSet(array( $this->marker ));
  }else{
    $this->bulletGraph->addDataSet(array( ($this->marker-$this->yourscore) ));
  }

  // green
  $this->bulletGraph->addValueMarkers('r', '60fb7a','0',$this->marker/$this->max,($this->marker/$this->max+0.004),'1');
  // yellow
  $this->bulletGraph->addValueMarkers('r', 'fff200', '0',$this->brokers/$this->max,($this->brokers/$this->max+0.008),'2');

  // score marker should be red for inclusive score range of 0-60
  if($this->yourscore <= 60){
    if( $this->yourscore == -1 ){
      $this->bulletGraph->setColors(array("a6a6a6","60fb7a"));
      $this->bulletGraph->addValueMarkers('r','a6a6a6','0',0/$this->max,(0/$this->max+0.008),'2');
    }else{
      $this->bulletGraph->setColors(array("ff0000","60fb7a"));
      $this->bulletGraph->addValueMarkers('r','ff0000','0',$this->yourscore/$this->max,($this->yourscore/$this->max+0.002),'2');
    }
  }else{
    $this->bulletGraph->setColors(array("000000","60fb7a"));
    $this->bulletGraph->addValueMarkers('r','000000','0',$this->yourscore/$this->max,($this->yourscore/$this->max+0.002),'2');
  }
    $this->bulletGraph->setBarWidth('7','15','15'); //sets up the width of the performance bar and the space either side.
    $this->bulletGraph->addAxisRange(0, 0, $this->max, 10); //sets the length of the x axis in terms of numbers of performance.
    $this->bulletGraph->setDataRange(0, $this->max); //sets the allowed range of data, in this case our sales figure will sit between 0 and the max length of the graph.

    /**
     * Next we setup the ranges poor, average and good again using range markers and the percentage calculation. We also set a series of gray scale colours for each range moving from dark to light tone.
     */
//poor
    $this->bulletGraph->addValueMarkers('r','0097db','0','0',$this->range1/$this->max,'0');
//average
    $this->bulletGraph->addValueMarkers('r','56bae7','0',$this->range1/$this->max,$this->range2/$this->max);
//good
    $this->bulletGraph->addValueMarkers('r','aadcf3','0',$this->range2/$this->max,'1');


    $this->bulletGraph->setHorizontal(true); //This sets the graph on its side as a bullet graph (If you require a vertical bullet graph you can turn this off and switch the visible axis from eariler).
    $this->bulletGraph->setVisibleAxes(array('x','y')); //display the axix label along the x axis i.e. the count of sales

    /**
     * Add a label to explain what the bullet graph is for...
     */
    $this->bulletGraph->addAxisLabel(1,array(str_replace(' ', '%20', $this->title)));

    /**
     * Return the Google Charts URL to generate the graph on your template, view or echo to the page wrapped in an <img> tag.
     */
    $this->imageURL = $this->bulletGraph->getUrl();
  }

  /**
   * @param mixed $brokers
   */
  public function setBrokers($brokers)
  {
    $this->brokers = $brokers;
  }

  /**
   * @param mixed $yourscore
   */
  public function setYourscore($yourscore)
  {
    $this->yourscore = $yourscore;
  }

  /**
   * @param mixed $marker
   */
  public function setMarker($marker)
  {
    $this->marker = $marker;
  }


  /**
   * @param mixed $range
   */
  public function setRange($range)
  {
    $this->range = $range;

    $this->range1 = $this->range[0];
    $this->range2 = $this->range[1];
    $this->max = $this->range[2];
  }

  /**
   * @return string
   */
  public function getImageURL()
  {
    return $this->imageURL;
  }
}
