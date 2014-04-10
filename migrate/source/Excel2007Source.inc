<?php
/**
 * @file
 *   Implementation of XLS Source for migration
 */

class Excel2007Source extends MigrateSource {

  /**
   * XLSX file path
   *
   * @var string
   */
  protected $file = '';

  /**
   * The number of rows in the CSV file before the data starts.
   *
   * @var integer
   */
  protected $headerRows = 0;

  /**
   * List of available source fields.
   *
   * @var array
   */
  protected $fields = array();

  /**
   * The current row/line number in the CSV file.
   *
   * @var integer
   */
  protected $rowNumber;

  /**
   * XLS Handler
   *
   * @var SimpleXLSX
   */
  protected $xlsDoc;

  /**
   * Column used to indentify empty rows (during count)
   *
   * @var string
   */
  protected $notEmptyColumn;

  /**
   * When this is true, we overwrite the values and if a column is empty, we'll
   * use the value for a previous row that has that column value not empty.
   *
   * @var bool
   */
  protected $reuseValues = FALSE;

  /**
   * New columns build by concatenating values from multiple columns
   *
   * @var array()
   */
  protected $concatColumns = array();

  /**
   * Alter callbacks for reused values
   *
   * If $reuseValues is set to FALSE, this callbacks will be ignored
   *
   * @see $reuseValues
   * @see alterReusedValues()
   */
  protected $reuseAlterCallbacks = array();

  /**
   * Raw line alter callbacks
   *
   * Functions called just after the line reading from xls document, after
   * reused values (if enabled), xls_line and concanated columns are added.
   *
   * Each function should have one parameter passed by reference and can do any
   * changes on it.
   *
   * Tipical use of this is to create custom source keys. For example when some
   * of the records use A column as source key and others use B. prepareRow() is
   * not called inside MigrateSource::next() and they can be skipped before
   * prepareRow() is called.
   *
   * This callbacks are called even if $reuseValues is set to FALSE.
   *
   * @see __construct()
   * @see rawLineAlter()
   */
  protected $rawLineAlterCallbacks = array();

  /**
   * Simple initialization.
   *
   * @param string $path
   *  The path to the source file
   * @param array $columns
   *  Keys are integers. values are array(field name, description).
   * @param array $options
   *  Options applied to this source.
   * @param array $fields
   *  Optional - keys are field names, values are descriptions. Use to override
   *  the default descriptions, or to add additional source fields which the
   *  migration will add via other means (e.g., prepareRow()).
   */
  public function __construct($path, array $columns = array(), array $options = array(), array $fields = array()) {
    parent::__construct($options);
    $this->file = $path;

    if (!empty($options['header_rows'])) {
      $this->headerRows = $options['header_rows'];
    }

    if (isset($options['not_empty_column'])) {
      $this->notEmptyColumn = $options['not_empty_column'];
    }

    if (isset($options['reuse_values'])) {
      $this->reuseValues = $options['reuse_values'];
    }
    if (isset($options['reuse_values_alter']) && is_array($options['reuse_values_alter'])) {
      $this->reuseAlterCallbacks = $options['reuse_values_alter'];
    }

    if (isset($options['raw_line_alter']) && is_array($options['raw_line_alter'])) {
      $this->rawLineAlterCallbacks = $options['raw_line_alter'];
    }

    $this->sheetNumber = 1; // By default use the first
    if (isset($options['sheet_number'])) {
      $this->sheetNumber = $options['sheet_number'];
    }

    $this->fields = $fields;

    // Create XLSX document handler
    if (!class_exists('SimpleXLSX')) {
      // This doesn't work inside test script. Investigate and maybe improve libraries support by
      // defining library inside this module
      //$library_path = module_invoke('libraries', 'get_path', 'simple-xlsx') . '/simplexlsx.class.php';
      $library_path = 'sites/all/libraries/simple-xlsx/simplexlsx.class.php';
      if (file_exists($library_path)) {
        include_once $library_path;
      }
      else {
        throw new Exception(t('SimpleXLSX library could not be found at: ' . $library_path . '!'));
      }
    }

    if (file_exists(drupal_realpath($path))) {
      $this->xlsDoc = new SimpleXLSX(drupal_realpath($path));
    }
    else {
      throw new Exception(t('Import file not found: ' . $path));
    }

    // Build the first line of the XLS or the last line defined as header
    $this->columns = array();
    $this->performRewind();
    $row = isset($this->rows[$this->headerRows - 1]) ? (array)$this->rows[$this->headerRows - 1] : reset($this->rows);

    foreach ($row as $header => $value) {
      $this->columns[$this->numericToColumn($header)] = "$value";
    }
    // Description passed as $columns argument will replace default description
    foreach($columns as $name => $description) {
      $this->columns[$name] = $description;
    }
  }

  /**
   * Return the schema used for a column to be used as key scheme
   *
   * The field is declared as varchar with maximum 255 chars
   *
   * @param string $column
   *   Column name uppercase
   *
   * @return array
   *   Key schema
   */
  public function getColumnSchema($column = 'A') {
    return array(
      $column => array(
        'type' => 'varchar',
        'length' => 255,
        'not null' => FALSE,
        'description' => "Column(s) '$column' from the Excel file",
      ),
    );
  }

  public function fields() {
    $fields = array();
    // the ones readed from file
    if (count($this->columns)) {
      foreach($this->columns as $name => $description) {
        $fields[$name] = $description;
      }
    }
    if (count($this->fields)) {
      foreach($this->fields as $name => $description) {
        $fields[$name] = $description;
      }
    }
    if (count($this->concatColumns)) {
      foreach($this->concatColumns as $field => $info) {
        $fields[$field] = t('Concatenation of columns: !columns',
                            array('!columns' => implode(', ', $info['columns'])));
      }
    }
    return $fields;
  }

  public function computeCount() {
    $empty_rows = 0; // This value will be substracted from the total
                     // along with rows declared as header

    try {
      $rows = $this->xlsDoc->rows($this->sheetNumber);
    }
    catch (Exception $e) {
      $rows = array();
    }
    if ($count = count($rows)) {
      // Remove empty rows
      foreach($rows as $i => $row) {
        if ($i < $this->headerRows) {
          continue;
        }

        if ($this->isEmptyRow($row)) {
          ++$empty_rows;
        }
      }
    }

    return $count - $this->headerRows - $empty_rows;
  }

  /**
   * String conversion
   *
   * Used as a cacheKey for Source::count() and displayed over the source
   * available fields as query.
   */
  public function __toString() {
    $parts = array();
    $parts[] = basename($this->file);
    $parts[] = round((int)filesize($this->file) / 1024) . 'KB';
    $parts[] = date("r", filemtime($this->file));
    return implode(' | ', $parts);
  }

  /**
   * Implementation of MigrateSource::performRewind().
   *
   * @return void
   */
  public function performRewind() {
    $this->rows = $this->xlsDoc->rows($this->sheetNumber);
    $this->rowNumber = $this->headerRows;
    $this->previousRow = new stdClass;
  }

  /**
   * Implementation of MigrateSource::getNextRow().
   * Return the next line of the source CSV file as an object.
   *
   * @return null|object
   */
  public function getNextRow() {
    if (isset($this->rows[$this->rowNumber])) {
      // Search for the next row that's not empty
      while ($this->isEmptyRow($this->rows[$this->rowNumber]) && isset($this->rows[$this->rowNumber + 1])) {
        $this->rowNumber++;
      }

      // If the founded row is empty, end this search
      if ($this->isEmptyRow($this->rows[$this->rowNumber])) {
        return NULL; // end this
      }

      $row = array();
      $max_col = max(array_keys($this->rows[$this->rowNumber]) + array(26));
      for ($i = 0; $i <= $max_col; ++$i) {
        $letter = $this->numericToColumn($i);
        if (isset($this->rows[$this->rowNumber][$i])) {
          $row[$letter] = $this->rows[$this->rowNumber][$i];
        }
      }

      $row = (object)$row;

      if ($this->reuseValues) {
        $old_values = $this->previousRow;

        // Give others a change to change the old values (use it for reset)
        $this->alterReusedValues($row, $old_values);

        // Overwrite missing fields with old values (if any)
        if (!empty($old_values)) {
          foreach($old_values as $col => $val) {
            if (!empty($val) && empty($row->{$col})) {
              $row->{$col} = $val;
            }
          }
        }

        // save the changes for the next row
        $this->previousRow = $row;
      }

      $row->xls_line = ++$this->rowNumber; // Line number as displayed in Excell (from 1 to N)
      $this->addConcatFields($row);

      // rawLineAlter functions
      $this->rawLineAlter($row);

      return $row;
    }
    return NULL;
  }

  /**
   * Transform column name to numeric index for a column considering A as column 0
   *
   * @param string $column
   *   Column name (eg: 'A', 'BC')
   *
   * @return int
   *   Index value (eg: A -> 0, C -> 2)
   */
  protected function columnToNumeric($column) {
    if (strlen($column) == 1) {
      // A = 1, B = 2
      return ord($column) - ord('A');
    }

    $last_char = substr($column, -1);
    $prefix = substr($column, 0, -1);
    return $this->columnIndex($last_char) +
      26 * ($this->columnIndex($prefix) + 1);
  }

  protected function numericToColumn($col) {
    static $n = 0;

    if (!$n) {
      $n = ord('Z') - ord('A') + 1;
    }

    if ($col < $n) {
      return chr($col + ord('A'));
    }

    return $this->numericToColumn(floor($col / $n) - 1) . $this->numericToColumn($col % $n);
  }

  /**
   * Check if an cell is empty
   */
  private function emptyCell($cell) {
    $c = trim($cell);
    return !empty($c);
  }

  /**
   * Check if a row in empty (has cells that are not empty)
   *
   * @param array $row
   *   Row values
   *
   * @return bool
   *   TRUE if the row contains only empty cells or FALSE otherwise
   */
  private function isEmptyRow($row) {
    return count(array_filter($row, array($this, 'emptyCell'))) == 0;
  }

  /**
   * Define a new concat column
   *
   * @param array $columns
   *   The name of the columns from xlsx file (eg. A, B, C)
   * @param string $field_name
   *   The name of the resulting field
   * @param string $value_separator
   *   What separator to use when concatenating values
   */
  public function createConcatColumn($columns, $field_name, $value_separator = '') {
    $this->concatColumns[$field_name] = array(
      'columns' => $columns,
      'separator' => $value_separator
    );
  }

  /**
   * Generate concat fields
   *
   * @param object $row
   *   Row object
   */
  protected function addConcatFields(&$row) {
    foreach ($this->concatColumns as $field_name => $info) {
      $values = array();
      foreach ($info['columns'] as $column) {
        $values[] = isset($row->{$column}) ? $row->{$column} : '';
      }
      $row->{$field_name} = implode($info['separator'], $values);
    }
  }

  /**
   * Pass old and new values to registered function
   *
   * Example: this functions can be used to reset reused values each time it
   * gets to a parent row by setting old_values to an empty stdClass object.
   *
   * @param object $row
   *   New row values
   * @param object $old_values
   *   Old values
   */
  protected function alterReusedValues(&$row, &$old_values) {
    foreach($this->reuseAlterCallbacks as $callback) {
      if (is_array($callback)) {
        if (method_exists($callback[0], $callback[1])) {
          if (is_string($callback[0])) {
            // call static method
            $callback[0]::$callback[1]($row, $old_values);
          }
          else {
            // call normal method on an object
            $callback[0]->{$callback[1]}($row, $old_values);
          }
        }
      }
      elseif (function_exists($callback)) {
        $callback($row, $old_values);
      }
    }
  }

  /**
   * Pass raw line to registered function
   *
   * @see $rawLineAlterCallbacks
   *
   * @param object $row
   *   New row values
   */
  protected function rawLineAlter(&$row) {
    // Clean up all migration added properties. This will avoid the reuse of
    // mapping attributes to be used for wrong records (souceid1, destid1 etc)
    foreach($row as $field_name => $value) {
      if (substr($field_name, 0, 8) == 'migrate_') {
        unset($row->{$field_name});
      }
    }

    // Call all callbacks
    foreach($this->rawLineAlterCallbacks as $callback) {
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

  /**
   * Add auxiliary fields after initialization
   *
   * Introduced to be used by ExpedProductMigration::addTaxonomyField().
   *
   * @param array $fields
   *   Array with field definition, key are field names and values are used
   *   as description
   */
  public function addExtraField($fields) {
    foreach($fields as $key => $name) {
      $this->fields[$key] = $name;
    }
  }

  // Debugging only
  public function getXLSDoc() {
    return $this->rows;
  }
}
