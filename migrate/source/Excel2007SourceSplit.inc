<?php
/**
 * @file
 *   XLSX Source plugin wich can split a row into multiple entries based on the
 *   values of one column
 */

class Excel2007SourceSplit extends XlsxSource {
  /**
   * Subvalue index
   *
   * If the selected rows has multiple values, this value indicates which one to
   * return
   *
   * @var int
   */
  private $subvalueIndex = 0;

  /**
   * Original row returned by parent class. Use it to iterate over all subvalues
   *
   * @var object
   */
  protected $original_row = NULL;

  /**
   * {@inheritdoc}
   */
  public function __construct($path, array $columns = array(), array $options = array(), array $fields = array()) {
    if (isset($options['raw_line_alter']) && is_array($options['raw_line_alter'])) {
      // These callbacks will be called from getNextRow by this class
      $this->rawLineAlterCallbacks2 = $options['raw_line_alter'];
      unset($options['raw_line_alter']);
    }

    parent::__construct($path, $columns, $options, $fields);

    // Do something with $options (save it)
    $this->options = $options;
  }

  // Reset subvalueIndex
  public function performRewind() {
    parent::performRewind();
    $this->subvalueIndex = 0;
    $this->original_row = NULL;
  }


  // Split into multiple rows
  public function getNextRow() {
    if ($this->original_row === NULL                                       // First call
      || (isset($this->original_row->subvalues_count)                      // Or we have record
          && $this->original_row->subvalues_count <= $this->subvalueIndex) // but we parsed all subvalues
    ) {
      $this->original_row = parent::getNextRow();
    }

    $return = NULL;

    if (!empty($this->original_row)) {
      // If the values are not splitted yet, clone the object and split the values
      if (!isset($this->original_row->source_values_splitted)) {
        $this->original_row = clone $this->original_row;

        // Parse columns and split the specified ones
        foreach($this->options['split_columns'] as $column => $separator) {
          if (isset($this->original_row->{$column})) {
            $this->original_row->{$column} = explode($separator, $this->original_row->{$column});
            $this->original_row->{$column} = array_map('trim', $this->original_row->{$column});
            $this->original_row->{$column} = array_filter($this->original_row->{$column});
            $this->original_row->{$column} = array_values($this->original_row->{$column}); // reset array keys
          }
          else {
            $this->original_row->{$column} = array();
          }

          // The first column that is split gives the total number of subvalues
          if (!isset($this->original_row->subvalues_count)) {
            $this->original_row->subvalues_count = count($this->original_row->{$column});
          }
        }

        $this->original_row->source_values_splitted = TRUE; // Next time will not split the values again
        $this->subvalueIndex = 0; // start from zero for each row
      }

      // Build the row for a given set of subvalues
      $return = clone $this->original_row;
      foreach($this->options['split_columns'] as $column => $separator) {
        $return->{$column} = isset($return->{$column}[$this->subvalueIndex])
                              ? $return->{$column}[$this->subvalueIndex]
                              : ''; // There is no subvalue for this index
      }
      $return->xls_subvalue_index = $this->subvalueIndex;
      $this->rawLineAlter2($return);

      // Increment the subvalue index
      $this->subvalueIndex++;
    } // end if not empty $this->orginal_row

    return $return;
  }


  // Disable count
  public function count($refresh = FALSE) {
    return -1;
  }

  /**
   * Pass raw line to registered function
   *
   * @see $rawLineAlterCallbacks
   *
   * @param object $row
   *   New row values
   */
  protected function rawLineAlter2(&$row) {
    // Call all callbacks
    foreach($this->rawLineAlterCallbacks2 as $callback) {
      if (is_array($callback)) {
        if (method_exists($callback[0], $callback[1])) {
          if (is_string($callback[0])) {
            // call static method
            $callback[0]::$callback[1]($row);
          }
          else {
            // call normal method on an object
            $callback[0]->{$callback[1]}($row);
          }
        }
      }
      elseif (function_exists($callback)) {
        $callback($row);
      }
    }
  }
}
