<?php

namespace customit\excelreport;

use app\models\AnketaSearch;
use Yii;
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Common\Exception\UnsupportedTypeException;
use Box\Spout\Common\Type;
use Box\Spout\Writer\Style\BorderBuilder;
use Box\Spout\Writer\Style\StyleBuilder;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Writer\Style\Border;
use Box\Spout\Writer\Style\Color;
use yii\base\InvalidConfigException;
use box\spout;
use yii\queue\Queue;

class ExcelReportModel {

    /*
     * @var ActiveDataProvider
     */
    private $_provider;
    /*
     * @var array
     */
    private $_columns;
    /*
     * @var string
     */
    public $filename;
    /*
     * @var bool
     */
    public $stripHtml = true;
    /*
     * @var string
     */
    public $searchClass;
    /*
     * @var string
     */
    public $folder = '@app/runtime/export';
    /*
     * @var array
     */
    public $headerStyleOptions = [];
    /*
     * @var array
     */
    public $boxStyleOptions = [];
    /*
     * @var bool
     */
    public $enableFormatter = true;
    /*
     * @var array
     */
    public $styleOptions = [];

    /*
     * @var WriterInterface
     */
    protected $_objWriter;

    protected $_objWorksheet;
    /*
     * @var int
     */
    protected $_endCol = 1;
    /*
     * @var string
     */
    protected $_exportType = 'xlsx';
    /*
     * @var int
     */
    protected $_endRow = 0;
    /*
     * @var int
     */
    protected $_beginRow = 1;
    /*
     * @var Queue
     */
    protected $queue;

    /**
     * ExcelReportModel constructor.
     * @param array $columns array of gridview columns
     * @param Queue $queue current queue for status reports
     * @param string $fileName
     * @param string $searchClass
     * @param string $searchMethod
     * @param array $searchParams
     * @param array $config
     */
    public function __construct($columns, $queue, $fileName, $searchClass, $searchMethod, $searchParams, array $config = [])
    {
        $searchModel = new $searchClass;
        $dataProvider = call_user_func_array([$searchModel, $searchMethod], $searchParams);
        $this->_provider = $dataProvider;
        $this->_columns = $columns;
        $this->queue = $queue;
        $this->filename = $fileName;
    }

    /**
     * Entry point
     */
    public function start() {
        $config = [
            'extension' => 'xlsx',
            'writer' => Type::XLSX,
        ];
        $this->initExport();
        try {
            $this->initExcelWriter($config);
        } catch (InvalidConfigException $e) {
            Yii::error($e->getMessage());
            exit;
        } catch (IOException $e) {
            Yii::error($e->getMessage());
            exit;
        } catch (UnsupportedTypeException $e) {
            Yii::error($e->getMessage());
            exit;
        }
        $this->initExcelWorksheet();
        $this->generateHeader();
        $this->generateBody();
        //Write data to file
        $this->_objWriter->close();
        //Unset vars
        $this->cleanup();
    }

    /**
     * Initializes export settings
     */
    public function initExport()
    {
        $this->setDefaultStyles('header');
        $this->setDefaultStyles('box');

        if (!isset($this->filename)) {
            $this->filename = 'grid-export';
        }
    }

    /**
     * Appends slash to path if it does not exist
     *
     * @param string $path
     * @param string $s the path separator
     *
     * @return string
     */
    public static function slash($path, $s = DIRECTORY_SEPARATOR)
    {
        $path = trim($path);
        if (substr($path, -1) !== $s) {
            $path .= $s;
        }
        return $path;
    }

    /**
     * Sets default styles
     *
     * @param string $section
     */
    protected function setDefaultStyles($section)
    {
        $defaultStyle = [];
        $opts = '';
        if ($section === 'header') {
            $opts = 'headerStyleOptions';

            $border = (new BorderBuilder())
                ->setBorderBottom(Color::BLACK, Border::WIDTH_MEDIUM, Border::STYLE_SOLID)
                ->build();
            $defaultStyle = (new StyleBuilder())
                ->setFontBold()
                ->setBackgroundColor('FFE5E5E5')
                ->setBorder($border)
                ->build();

        } elseif ($section === 'box') {
            $opts = 'boxStyleOptions';

            $border = (new BorderBuilder())
                ->setBorderBottom(Color::BLACK, Border::WIDTH_MEDIUM, Border::STYLE_SOLID)
                ->build();
            $defaultStyle = (new StyleBuilder())
                ->setBorder($border)
                ->build();
        }
        if (empty($opts)) {
            return;
        }

        $this->$opts = $defaultStyle;
    }

    /**
     * Initializes Spout Writer Object Instance
     *
     * @param array $config
     * @throws InvalidConfigException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Common\Exception\IOException
     */
    public function initExcelWriter($config)
    {
        $this->folder = trim(Yii::getAlias($this->folder));
        $file = self::slash($this->folder) . $this->filename . '.' . $config['extension'];
        if (!file_exists($this->folder) && !mkdir($this->folder, 0777, true)) {
            throw new InvalidConfigException(
                "Invalid permissions to write to '{$this->folder}' as set in `Export::folder` property."
            );
        }

        $this->_objWriter = WriterFactory::create($config['writer']);
        $this->_objWriter->setShouldUseInlineStrings(true);
        $this->_objWriter->setShouldUseInlineStrings(false);

        $this->_objWriter->openToFile($file);
    }

    /**
     * Get Worksheet Instance
     */
    public function initExcelWorksheet()
    {
        $this->_objWorksheet = $this->_objWriter->getCurrentSheet();
        $this->_objWorksheet->setName('Выгрузка');
    }

    /**
     * Returns an excel column name.
     *
     * @param integer $index the column index number
     *
     * @return string
     */
    public static function columnName($index)
    {
        $i = intval($index) - 1;
        if ($i >= 0 && $i < 26) {
            return chr(ord('A') + $i);
        }
        if ($i > 25) {
            return (self::columnName($i / 26)) . (self::columnName($i % 26 + 1));
        }
        return 'A';
    }

    /**
     * Generates the output data header content.
     */
    public function generateHeader()
    {
        if (count($this->_columns) == 0) {
            return;
        }
        $styleOpts = $this->headerStyleOptions;
        $headValues = [];
        $this->_endCol = 0;
        //Generate labels array
        foreach ($this->_columns as $column) {
            $this->_endCol++;
            $head = (isset($column['label'])) ? $column['label'] : '#';
            $headValues[] = $head;
        }
        //Write header content
        $this->setRowValues($headValues, $styleOpts);
    }

    /**
     * Sets the values of excel row
     *
     * @param array $values
     * @param array $style
     *
     */
    protected function setRowValues($values, $style = null)
    {
        if ($this->stripHtml) {
            array_map('strip_tags', $values);
        }
        array_walk($values, function (&$item, $key) {
            html_entity_decode($item, ENT_QUOTES, 'UTF-8');
        });

        if (!empty($style)) {
            $this->_objWriter->addRowWithStyle($values, $style);
        } else {
            $this->_objWriter->addRow($values);
        }
    }

    /**
     * Generates the output data body content.
     *
     * @return integer the number of output rows.
     */
    public function generateBody()
    {
        $this->_endRow = 0;
        $offset = 1000;
        $totalCount = $this->_provider->getTotalCount();
        $loops = ceil($totalCount / $offset);

        for ($i=0; $i <= $loops; $i++) {
            $data = $this->_provider->query->limit($offset)->offset($i*$offset)->all();
            foreach ($data as $key => $value) {
                $this->generateRow($value, $this->_endRow);
                $this->_endRow++;
                //Change queue process progress
                $this->queue->setProgress($this->_endRow, $totalCount);
            }
        }

        return $this->_endRow;
    }

    /**
     * Generates an output data row with the given data.
     *
     * @param mixed $data the data model to be rendered
     */
    public function generateRow($data)
    {
        $this->_endCol = 0;
        $rowData = [];
        foreach ($this->_columns as $column) {
            if (!empty($column['hiddenFromExport'])) {
                continue;
            }
            $format = $this->enableFormatter && isset($column['format']) ? $column['format'] : 'raw';
            $value = null;
            if (isset($column['class']) && $column['class'] == 'yii\\grid\\ActionColumn') {
                $value = null;
            } elseif (isset($column['class']) && $column['class'] == 'yii\\grid\\SerialColumn') {
                $value = $this->_endRow+1;
            } elseif (isset($column['value'])) {
                $var = $column['value'];
                $value = $data->$var;
            } elseif (isset($column['attribute'])) {
                $var = $column['attribute'];
                $value = $data->$var;
            }
            $this->_endCol++;
            if (isset($value) && $value !== '' && isset($format)) {
                $value = Yii::$app->formatter->format($value, $format);
            } else {
                $value = '';
            }
            $rowData[] = $value;
        }

        $this->setRowValues($rowData);
    }

    /**
     * Cleans up the export file and current object instance
     */
    protected function cleanup()
    {
        unset($this->_provider, $this->_objWriter);
    }
}
